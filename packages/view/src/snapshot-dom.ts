/**
 * 单页快照的浏览器那一半：收宿主的 `@font-face`、把独立 SVG 画上 canvas、出 PNG。
 * 为什么不截屏幕、为什么要内联，见 `snapshot.ts` 的文件头。
 */
import { twipsToPx } from '@uw/core';
import type { PageLayout } from '@uw/layout';
import type { MountOptions } from '@uw/render-dom/dom';
import type { FetchBytes, FontFaceSource } from './snapshot.ts';
import { normalizeFamily, pageSnapshotMarkup } from './snapshot.ts';

export interface PngOptions {
  /**
   * 导出倍率：1 = 文档在 100% 缩放下的 CSS px（96 dpi），2 = 192 dpi。默认 1。
   * **与视图当前的屏幕缩放无关** —— 与打印同理，导出只认纸张。
   */
  scale?: number;
}

/** canvas 单边上限：Chrome / Firefox 约 32767px，Safari 更小；超了 toBlob 答 null 而不报错，只能事先拦 */
const MAX_CANVAS_SIDE = 16384;

/**
 * 收集宿主页面上的 `@font-face`。跨域样式表读 `cssRules` 会抛 SecurityError，只能跳过 ——
 * 那种字体在快照里退到本机字体；`document.fonts.add(new FontFace(…))` 加进来的字体
 * 拿不到字节，同样退到本机字体。
 */
export function collectFontFaces(doc: Document): FontFaceSource[] {
  const out: FontFaceSource[] = [];
  const seen = new Set<CSSStyleSheet>();
  // 按 `rule.type` 认而不是 `instanceof win.CSSFontFaceRule`：这几个常量虽标了过时，各引擎都在，
  // 而规则类不是每个环境都挂在 window 上（jsdom 就没有 CSSFontFaceRule）
  const FONT_FACE = 5;
  const IMPORT = 3;
  const walk = (rules: CSSRuleList, baseUrl: string) => {
    for (const rule of Array.from(rules)) {
      if (rule.type === FONT_FACE) {
        const family = normalizeFamily((rule as CSSFontFaceRule).style.getPropertyValue('font-family'));
        if (family !== '') out.push({ family, cssText: rule.cssText, baseUrl });
      } else if (rule.type === IMPORT) {
        const imported = (rule as CSSImportRule).styleSheet;
        if (imported !== null) sheet(imported);
      } else if ('cssRules' in rule) {
        // @media / @supports / @layer：里面的 @font-face 照样生效
        walk((rule as CSSGroupingRule).cssRules, baseUrl);
      }
    }
  };
  const sheet = (s: CSSStyleSheet) => {
    if (seen.has(s)) return;
    seen.add(s);
    let rules: CSSRuleList;
    try {
      rules = s.cssRules;
    } catch {
      return;
    }
    walk(rules, s.href ?? doc.baseURI);
  };
  for (const s of Array.from(doc.styleSheets)) sheet(s);
  for (const s of doc.adoptedStyleSheets ?? []) sheet(s);
  return out;
}

function fetcherOf(doc: Document): FetchBytes | undefined {
  const fetch = doc.defaultView?.fetch;
  if (fetch === undefined) return undefined;
  return async (url) => {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`${response.status} ${url}`);
    return {
      bytes: new Uint8Array(await response.arrayBuffer()),
      type: response.headers.get('content-type')?.split(';')[0]?.trim() ?? '',
    };
  };
}

async function loadImage(doc: Document, src: string): Promise<HTMLImageElement> {
  const img = doc.createElement('img');
  const loaded = new Promise<void>((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error('页面快照的 SVG 加载失败'));
  });
  img.src = src;
  await loaded;
  // onload 之后内联字体可能还在解码，decode() 等它真能画
  await img.decode().catch(() => {});
  return img;
}

/** 一页 → PNG。`opts` 是视图的渲染选项（字体映射、图片地址、class 前缀），屏幕缩放在这里被 `scale` 顶掉 */
export async function pageToPng(page: PageLayout, opts: MountOptions, png: PngOptions = {}): Promise<Blob> {
  const scale = png.scale ?? 1;
  if (!Number.isFinite(scale) || scale <= 0) throw new RangeError('scale 必须是大于 0 的有限数');
  const doc = opts.document ?? globalThis.document;
  const width = Math.round(twipsToPx(page.geometry.width, scale));
  const height = Math.round(twipsToPx(page.geometry.height, scale));
  if (width > MAX_CANVAS_SIDE || height > MAX_CANVAS_SIDE)
    throw new RangeError(`导出尺寸 ${width}×${height}px 超过画布上限 ${MAX_CANVAS_SIDE}px，调小 scale`);

  const { document: _, debug: __, ...render } = opts;
  const fetchBytes = fetcherOf(doc);
  const markup = await pageSnapshotMarkup(page, {
    ...render,
    zoom: scale,
    fontFaces: collectFontFaces(doc),
    ...(fetchBytes === undefined ? {} : { fetchBytes }),
  });
  // data URI 而不是 blob URL：部分浏览器把 blob URL 的 SVG 图算作跨源，画上去 canvas 就被污染、toBlob 抛错
  const img = await loadImage(doc, `data:image/svg+xml;charset=utf-8,${encodeURIComponent(markup)}`);
  const canvas = doc.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');
  if (ctx === null) throw new Error('当前环境不支持 canvas 2D');
  // 先铺白：纸宽折成 px 带小数（A4 = 793.73px），取整后的画布边上那一列只被盖住一部分，
  // 不铺底就是一圈半透明的边
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, width, height);
  ctx.drawImage(img, 0, 0, width, height);
  return new Promise<Blob>((resolve, reject) => {
    canvas.toBlob((blob) => (blob === null ? reject(new Error('PNG 编码失败')) : resolve(blob)), 'image/png');
  });
}
