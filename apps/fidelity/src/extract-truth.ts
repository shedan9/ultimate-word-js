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
import type { TruthFont, TruthItem, TruthLine, TruthPage, WordMeta, WordTruth } from './truth-types.ts';

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
    // 否则 TextItem.fontName 只能拿到 "g_d0_f1" 这类内部名
    try {
      await page.getOperatorList();
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
