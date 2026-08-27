/**
 * PDF → 坐标真值。
 *
 * pdf.js 给出的 transform 是 PDF 用户空间矩阵（原点左下、y 向上）；这里统一翻成
 * 左上原点、y 向下，并按基线聚合成行。输出要提交进仓库，所以刻意做成确定性的：
 * 不写时间戳，所有浮点固定截断到 3 位小数。
 */
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import type {
  TruthFont,
  TruthImage,
  TruthItem,
  TruthLine,
  TruthPage,
  TruthRule,
  WordMeta,
  WordTruth,
} from './truth-types.ts';

const require = createRequire(import.meta.url);
const PDFJS_ROOT = path.dirname(require.resolve('pdfjs-dist/package.json'));
const PDFJS_VERSION: string = require('pdfjs-dist/package.json').version;

/** pdf.js 的资源目录必须以 URL 形式、且以 "/" 结尾传入 */
const asDirUrl = (...seg: string[]) => `${pathToFileURL(path.join(PDFJS_ROOT, ...seg)).href}/`;

export interface ExtractOptions {
  /** 同一行的基线容差（pt）。Word 导出的 PDF 里同行基线通常完全相等，留一点余量给上下标之外的微差 */
  lineTolerance?: number;
  /** Word COM 导出的 sidecar 元数据，用于交叉校验页数与页面设置 */
  meta?: WordMeta | undefined;
  /** 记进 truth.json 的源文件名；默认取 PDF 的 basename 换成 .docx */
  source?: string;
}

const r3 = (v: number): number => {
  const r = Math.round(v * 1000) / 1000;
  return Object.is(r, -0) ? 0 : r;
};

/** 取首/末**码点**，避免代理对被切半 */
const firstCp = (s: string): string => [...s][0] ?? '';
const lastCp = (s: string): string => {
  const cps = [...s];
  return cps[cps.length - 1] ?? '';
};

export async function extractTruth(pdfPath: string, opts: ExtractOptions = {}): Promise<WordTruth> {
  const tol = opts.lineTolerance ?? 0.6;
  // legacy 构建是官方给 Node 用的那份（不依赖 DOM）
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');

  const data = new Uint8Array(await readFile(pdfPath));
  const loadingTask = pdfjs.getDocument({
    data,
    // 只抽坐标不画字：关掉字体实例化，省掉一堆 Node 下的 fontFace 噪声
    disableFontFace: true,
    useSystemFonts: false,
    cMapUrl: asDirUrl('cmaps'),
    cMapPacked: true,
    standardFontDataUrl: asDirUrl('standard_fonts'),
    wasmUrl: asDirUrl('wasm'),
  });
  const doc = await loadingTask.promise;

  const pages: TruthPage[] = [];
  const fonts: Record<string, TruthFont> = {};
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const viewport = page.getViewport({ scale: 1 });
    const pageHeight = viewport.height;

    // 先跑一遍算子表：字体对象要等它跑完才进 commonObjs，
    // 否则 TextItem.fontName 只能拿到 "g_d0_f1" 这类内部名。
    // 图片与格线也只有这一路能拿到 —— getTextContent() 吐的是 show-text 的产物，
    // 画出来的东西不在里面
    let images: TruthImage[] = [];
    let rules: TruthRule[] = [];
    try {
      ({ images, rules } = collectPaint(await page.getOperatorList(), pageHeight, pdfjs.OPS));
    } catch {
      // 字体名拿不到不影响坐标真值，降级即可
    }

    const tc = await page.getTextContent({ disableNormalization: true });
    const resolve = fontResolver(page, tc.styles, fonts);
    const items: TruthItem[] = [];
    for (const raw of tc.items) {
      if (!('str' in raw)) continue; // marked-content 标记项
      const it = raw as { str: string; transform: number[]; width: number; fontName: string };
      if (it.str === '') continue;
      const t = it.transform;
      const [, , c = 0, d = 0, e = 0, f = 0] = t;
      items.push({
        x: r3(e),
        y: r3(pageHeight - f),
        w: r3(it.width),
        // 字号 = 文本矩阵纵向缩放的模，与 pdf.js textLayer 的算法一致
        size: r3(Math.hypot(c, d)),
        font: resolve(it.fontName),
        text: it.str,
      });
    }

    pages.push({
      index: i - 1,
      width: r3(viewport.width),
      height: r3(viewport.height),
      rotate: page.rotate,
      items,
      lines: groupLines(items, tol),
      ...(images.length > 0 ? { images } : {}),
      ...(rules.length > 0 ? { rules } : {}),
    });
    page.cleanup();
  }
  // destroy() 挂在 loadingTask 上，同时关掉 worker
  await loadingTask.destroy();

  const truth: WordTruth = {
    source: opts.source ?? `${path.basename(pdfPath, path.extname(pdfPath))}.docx`,
    generator: {
      tool: '@uw/fidelity',
      formatVersion: 1,
      pdfjs: PDFJS_VERSION,
      ...(opts.meta ? { word: opts.meta.wordVersion, wordBuild: opts.meta.wordBuild } : {}),
    },
    unit: 'pt',
    origin: 'top-left',
    pageCount: doc.numPages,
    fonts: Object.fromEntries(Object.entries(fonts).sort(([a], [b]) => (a < b ? -1 : 1))),
    ...(opts.meta ? { wordPageCount: opts.meta.pageCount, sections: opts.meta.sections } : {}),
    pages,
  };
  return truth;
}

/** 子集前缀（"BCDEEE+"）每次导出都不同，必须剥掉，否则真值 diff 全是噪声 */
const stripSubsetTag = (name: string): string => name.replace(/^[A-Z]{6}\+/, '');

/**
 * pdf.js 只在 TextItem 上给内部名（g_d0_f1）；真实名要拿这个 key 回查 commonObjs。
 * v6 起 commonObjs 的内部存储是私有字段，枚举不了，只能按 key 逐个取。
 */
function fontResolver(
  page: { commonObjs: { has(k: string): boolean; get(k: string): unknown } },
  styles: Record<string, { ascent?: number; descent?: number; vertical?: boolean }>,
  sink: Record<string, TruthFont>,
): (internalName: string) => string {
  const cache = new Map<string, string>();
  return (internalName: string): string => {
    const hit = cache.get(internalName);
    if (hit !== undefined) return hit;

    let name = internalName;
    try {
      if (page.commonObjs.has(internalName)) {
        const obj = page.commonObjs.get(internalName) as { name?: string } | null;
        if (obj && typeof obj.name === 'string') name = stripSubsetTag(obj.name);
      }
    } catch {
      // 解析不出就退回内部名，坐标真值不受影响
    }

    const style = styles[internalName];
    if (style && sink[name] === undefined) {
      sink[name] = {
        name,
        ascent: r3(style.ascent ?? 0),
        descent: r3(style.descent ?? 0),
        vertical: style.vertical === true,
      };
    }
    cache.set(internalName, name);
    return name;
  };
}

/**
 * 从算子表里读出**画出来的东西**：图片的落点与线（含底纹色块）。
 *
 * 两样合在一趟里演，是因为两者都只能靠 CTM 定位，而 CTM 要照着 `q` / `Q` / `cm`
 * 一路演下来 —— 演两遍就有两处会漏演 `q` 的机会。
 *
 * PDF 里图片**没有自己的坐标**：`Do` 把图片铺满**当前变换矩阵下的单位正方形**，
 * 位置与大小全在 CTM 里。少演一个 `q`，后面所有图片的坐标都会偏。
 * Form XObject（Word 导出的图常裹一层）自带一个矩阵，进出各是一次隐式的 `q` / `Q`，
 * 与 `save` / `restore` 一样要压栈。
 *
 * 图片结果取**外接矩形**并翻成左上原点：旋转过的图在 PDF 里是斜的，而我们要比的是
 * `wp:extent` 那个正的外框。
 *
 * 线这一路收两种：**填充**路径（`f` / `f*` / `B`，Word 的实线格线是一个个细长的填充矩形）
 * 与**描边**路径（`S`，虚线 / 点线走这条，dash 模式带在 `setDash` 上）。
 * 描边换算成它盖住的那个矩形（线心 ± 半个线宽），与填充那一路同形，比对代码只认一种数据；
 * 只收**轴对齐的单段直线**，斜线与折线跳过。
 * 裁剪路径（`W n`，pdf.js 里是 `constructPath` 的类型为 `endPath`）不算画出来的线，跳过：
 * Word 给**每个单元格**都下一个裁剪框，收进来的话每张表都会多出一圈假的外框。
 */
function collectPaint(
  opList: { fnArray: ArrayLike<number>; argsArray: unknown[] },
  pageHeight: number,
  OPS: Record<string, number>,
): { images: TruthImage[]; rules: TruthRule[] } {
  const images: TruthImage[] = [];
  const rules: TruthRule[] = [];
  let ctm: Matrix = [1, 0, 0, 1, 0, 0];
  let gs: GState = { fill: '#000000', stroke: '#000000', lineWidth: 1, dash: [] };
  // 颜色 / 线宽 / dash 都属于图形状态，`q` / `Q` 要连它们一起存取 —— 只压 CTM 的话，
  // 一个 `q` 里改过的颜色会漏到后面所有的线上
  const stack: { ctm: Matrix; gs: GState }[] = [];
  const paint = new Set([
    OPS.paintImageXObject,
    OPS.paintImageMaskXObject,
    OPS.paintInlineImageXObject,
    OPS.paintImageXObjectRepeat,
  ]);
  const filled = new Set([
    OPS.fill,
    OPS.eoFill,
    OPS.fillStroke,
    OPS.eoFillStroke,
    OPS.closeFillStroke,
    OPS.closeEOFillStroke,
  ]);
  const stroked = new Set([
    OPS.stroke,
    OPS.closeStroke,
    OPS.fillStroke,
    OPS.eoFillStroke,
    OPS.closeFillStroke,
    OPS.closeEOFillStroke,
  ]);

  for (let i = 0; i < opList.fnArray.length; i++) {
    const fn = opList.fnArray[i] as number;
    const args = opList.argsArray[i] as unknown[] | null;
    if (fn === OPS.save) {
      stack.push({ ctm, gs });
    } else if (fn === OPS.restore) {
      const st = stack.pop();
      ctm = st?.ctm ?? [1, 0, 0, 1, 0, 0];
      gs = st?.gs ?? INITIAL_GSTATE;
    } else if (fn === OPS.transform) {
      ctm = mul(ctm, args as unknown as Matrix);
    } else if (fn === OPS.paintFormXObjectBegin) {
      stack.push({ ctm, gs });
      const m = (args as unknown[] | null)?.[0];
      if (Array.isArray(m)) ctm = mul(ctm, m as unknown as Matrix);
    } else if (fn === OPS.paintFormXObjectEnd) {
      const st = stack.pop();
      ctm = st?.ctm ?? [1, 0, 0, 1, 0, 0];
      gs = st?.gs ?? INITIAL_GSTATE;
    } else if (fn === OPS.setFillRGBColor) {
      if (typeof args?.[0] === 'string') gs = { ...gs, fill: args[0] as string };
    } else if (fn === OPS.setStrokeRGBColor) {
      if (typeof args?.[0] === 'string') gs = { ...gs, stroke: args[0] as string };
    } else if (fn === OPS.setLineWidth) {
      if (typeof args?.[0] === 'number') gs = { ...gs, lineWidth: args[0] as number };
    } else if (fn === OPS.setDash) {
      const pattern = args?.[0];
      gs = { ...gs, dash: Array.isArray(pattern) ? (pattern as number[]) : [] };
    } else if (paint.has(fn)) {
      const name = typeof args?.[0] === 'string' ? (args[0] as string) : '';
      images.push(unitSquareRect(ctm, pageHeight, name));
    } else if (fn === OPS.constructPath) {
      const type = args?.[0] as number;
      if (filled.has(type)) {
        for (const rect of pathRects(args?.[1], ctm, pageHeight, gs.fill)) rules.push(rect);
      }
      if (stroked.has(type)) {
        for (const rect of strokeRects(args?.[1], ctm, pageHeight, gs)) rules.push(rect);
      }
    }
  }
  return { images, rules: dedupeRules(rules) };
}

/**
 * 一条填充路径 → 若干个矩形。
 *
 * pdf.js v6 把路径压成扁平数组：`[cmd, …坐标, cmd, …]`，cmd 见 `DrawOPS`
 * （0 moveTo / 1 lineTo / 2 curveTo / 3 quadraticCurveTo / 4 closePath）。
 * 我们只认由直线围成的**轴对齐矩形**：格线就是这一种，而曲线路径（圆角、艺术字）
 * 收进来只会变成一堆对不上的外接矩形。每个 `moveTo` 开一条新子路径，
 * 一次 `f` 可能填好几条 —— 所以是「返回若干个」而不是一个。
 */
function pathRects(pathData: unknown, ctm: Matrix, pageHeight: number, color: string): TruthRule[] {
  const out: TruthRule[] = [];
  const segs: ArrayLike<number>[] = Array.isArray(pathData)
    ? (pathData as ArrayLike<number>[])
    : [pathData as ArrayLike<number>];

  for (const seg of segs) {
    if (seg === undefined || seg === null || typeof seg.length !== 'number') continue;
    let pts: [number, number][] = [];
    let curved = false;
    const flush = (): void => {
      const rect = axisAlignedRect(pts, ctm, pageHeight, color);
      if (!curved && rect !== undefined) out.push(rect);
      pts = [];
    };
    for (let k = 0; k < seg.length; ) {
      const cmd = seg[k] as number;
      if (cmd === 0) {
        flush();
        pts.push([seg[k + 1] as number, seg[k + 2] as number]);
        k += 3;
      } else if (cmd === 1) {
        pts.push([seg[k + 1] as number, seg[k + 2] as number]);
        k += 3;
      } else if (cmd === 2) {
        curved = true;
        k += 7;
      } else if (cmd === 3) {
        curved = true;
        k += 5;
      } else {
        k += 1; // closePath：矩形本来就按点集判，闭不闭合不影响
      }
    }
    flush();
  }
  return out;
}

/** 点集是不是一个轴对齐矩形（允许闭合时重复首点）；是就翻成左上原点的矩形 */
function axisAlignedRect(
  pts: readonly [number, number][],
  ctm: Matrix,
  pageHeight: number,
  color: string,
): TruthRule | undefined {
  if (pts.length < 4 || pts.length > 5) return undefined;
  const dev = pts.map(([u, v]) => [ctm[0] * u + ctm[2] * v + ctm[4], ctm[1] * u + ctm[3] * v + ctm[5]]);
  const xs = dev.map((p) => p[0] as number);
  const ys = dev.map((p) => p[1] as number);
  const x0 = Math.min(...xs);
  const x1 = Math.max(...xs);
  const y0 = Math.min(...ys);
  const y1 = Math.max(...ys);
  // 每个点都必须落在外接矩形的角上，否则它是个多边形（Word 画斜线也走填充路径）
  const EPS = 1e-3;
  const onCorner = dev.every(
    (p) =>
      (Math.abs((p[0] as number) - x0) < EPS || Math.abs((p[0] as number) - x1) < EPS) &&
      (Math.abs((p[1] as number) - y0) < EPS || Math.abs((p[1] as number) - y1) < EPS),
  );
  if (!onCorner) return undefined;
  if (x1 - x0 < EPS && y1 - y0 < EPS) return undefined; // 退化成点
  // PDF 的 y 向上：矩形的**顶**边是 y 最大的那条
  return { x: r3(x0), y: r3(pageHeight - y1), w: r3(x1 - x0), h: r3(y1 - y0), color };
}

/** 图形状态里与「画出来的线长什么样」有关的那几项 */
interface GState {
  fill: string;
  stroke: string;
  lineWidth: number;
  dash: number[];
}

const INITIAL_GSTATE: GState = { fill: '#000000', stroke: '#000000', lineWidth: 1, dash: [] };

/**
 * 一条描边路径 → 它盖住的矩形。
 *
 * 只认**两个点的轴对齐直线段**：虚线 / 点线的边框就是这一种（dash 模式在图形状态里，
 * 不在路径上），别的形状（折线、斜线、曲线）不是格线，收进来只会变成对不上的外接矩形。
 * 线宽要过一遍 CTM —— Word 这里的 CTM 是单位阵，但一旦有缩放，`w` 是用户空间的数。
 * 端点样式（`J`）没有消费：Word 的格线用默认的平头端，两端不外扩。
 */
function strokeRects(pathData: unknown, ctm: Matrix, pageHeight: number, gs: GState): TruthRule[] {
  const out: TruthRule[] = [];
  const segs: ArrayLike<number>[] = Array.isArray(pathData)
    ? (pathData as ArrayLike<number>[])
    : [pathData as ArrayLike<number>];
  // 均匀缩放下线宽的换算系数；非均匀缩放（Word 不产出）时取行列式的开方作近似
  const scale = Math.sqrt(Math.abs(ctm[0] * ctm[3] - ctm[1] * ctm[2])) || 1;
  const half = (gs.lineWidth * scale) / 2;
  const EPS = 1e-3;

  for (const seg of segs) {
    if (seg === undefined || seg === null || typeof seg.length !== 'number') continue;
    const pts: [number, number][] = [];
    let ok = true;
    for (let k = 0; k < seg.length; ) {
      const cmd = seg[k] as number;
      if (cmd === 0 || cmd === 1) {
        pts.push([seg[k + 1] as number, seg[k + 2] as number]);
        k += 3;
      } else if (cmd === 4) {
        k += 1;
      } else {
        ok = false;
        break;
      }
    }
    if (!ok || pts.length !== 2) continue;
    const [p0, p1] = pts as [[number, number], [number, number]];
    const ax = ctm[0] * p0[0] + ctm[2] * p0[1] + ctm[4];
    const ay = ctm[1] * p0[0] + ctm[3] * p0[1] + ctm[5];
    const bx = ctm[0] * p1[0] + ctm[2] * p1[1] + ctm[4];
    const by = ctm[1] * p1[0] + ctm[3] * p1[1] + ctm[5];
    const vertical = Math.abs(ax - bx) < EPS;
    const horizontal = Math.abs(ay - by) < EPS;
    if (!vertical && !horizontal) continue;
    const x0 = Math.min(ax, bx) - (vertical ? half : 0);
    const x1 = Math.max(ax, bx) + (vertical ? half : 0);
    const y0 = Math.min(ay, by) - (horizontal ? half : 0);
    const y1 = Math.max(ay, by) + (horizontal ? half : 0);
    out.push({
      x: r3(x0),
      y: r3(pageHeight - y1),
      w: r3(x1 - x0),
      h: r3(y1 - y0),
      color: gs.stroke,
      ...(gs.dash.length > 0 ? { dash: gs.dash.map(r3) } : {}),
    });
  }
  return out;
}

/**
 * 去掉完全重合的重复矩形。Word 在每个格线交点上画一个小方块，相邻两格各画一遍，
 * 于是同一个方块常常出现两次 —— 留着只会让「这条边画了几段」这种计数全部翻倍。
 */
function dedupeRules(rules: readonly TruthRule[]): TruthRule[] {
  const seen = new Set<string>();
  const out: TruthRule[] = [];
  for (const r of rules) {
    const key = `${r.x},${r.y},${r.w},${r.h},${r.color},${r.dash?.join('/') ?? ''}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(r);
  }
  return out;
}

type Matrix = [number, number, number, number, number, number];

/** `m1` 之上再叠一个 `m2`（PDF 的 `cm` 语义：新矩阵先作用，再套原来的） */
function mul(m1: Matrix, m2: Matrix): Matrix {
  return [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
  ];
}

function unitSquareRect(m: Matrix, pageHeight: number, name: string): TruthImage {
  const xs: number[] = [];
  const ys: number[] = [];
  for (const [u, v] of [
    [0, 0],
    [1, 0],
    [1, 1],
    [0, 1],
  ] as const) {
    xs.push(m[0] * u + m[2] * v + m[4]);
    ys.push(m[1] * u + m[3] * v + m[5]);
  }
  const x = Math.min(...xs);
  const w = Math.max(...xs) - x;
  // PDF 的 y 向上：图的**顶**边是 y 最大的那条
  const top = pageHeight - Math.max(...ys);
  const h = Math.max(...ys) - Math.min(...ys);
  return { x: r3(x), y: r3(top), w: r3(w), h: r3(h), yBottom: r3(top + h), name };
}

/** 按基线把片段聚合成行；同基线内按 x 排序，得到稳定的行首/行末字符 */
export function groupLines(items: TruthItem[], tol: number): TruthLine[] {
  const order = items
    .map((_, i) => i)
    .sort((a, b) => {
      const ia = items[a] as TruthItem;
      const ib = items[b] as TruthItem;
      return ia.y - ib.y || ia.x - ib.x;
    });

  const lines: TruthLine[] = [];
  let cur: number[] = [];
  let curY = Number.NaN;

  const flush = () => {
    if (cur.length === 0) return;
    const idx = cur.slice().sort((a, b) => (items[a] as TruthItem).x - (items[b] as TruthItem).x);
    const parts = idx.map((i) => (items[i] as TruthItem).text);
    const text = parts.join('');
    const firstIt = items[idx[0] as number] as TruthItem;
    const lastIt = items[idx[idx.length - 1] as number] as TruthItem;
    lines.push({
      y: r3(curY),
      x: r3(firstIt.x),
      xEnd: r3(lastIt.x + lastIt.w),
      text,
      first: firstCp(text),
      last: lastCp(text),
      items: idx,
    });
    cur = [];
  };

  for (const i of order) {
    const it = items[i] as TruthItem;
    if (cur.length === 0) {
      curY = it.y;
    } else if (Math.abs(it.y - curY) > tol) {
      flush();
      curY = it.y;
    }
    cur.push(i);
  }
  flush();
  return lines;
}

// 允许直接 `node src/extract-truth.ts <pdf>` 单独抽一份真值
if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  const pdf = process.argv[2];
  if (!pdf) {
    console.error('usage: node src/extract-truth.ts <file.pdf>');
    process.exit(2);
  }
  const truth = await extractTruth(pdf);
  process.stdout.write(`${JSON.stringify(truth, null, 2)}\n`);
}
