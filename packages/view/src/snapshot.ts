/**
 * 单页快照 —— `view.toPNG(page)`（api.md §13）的纯数据那一半。
 *
 * 走的路是「同一份 `PageLayout` → 独立的 SVG 文本 → `<img>` → canvas」，不是截屏幕：
 * 屏幕上那一份被虚拟化卸掉了大半页、带着屏幕缩放，还可能压着装饰与 overlay。
 * 与打印同理，快照按 `scale` 重画一遍，**不重排**。
 *
 * 难处全在「SVG 当图片用」的沙箱：`<img>` 里的 SVG **不许加载任何外部资源**，
 * 也看不见宿主页面的 `@font-face`（本机装的字体倒是看得见）。所以两样东西要内联：
 * ① 宿主用 `@font-face` 引进来的网络字体 —— 只挑这一页真用到的族，一份仿宋 webfont 动辄 10 MB，
 *   全塞进去每导一页都要编码一遍；② 不是 data URI 的图片地址（宿主自己给的 blob URL / CDN 链接）。
 * 取字节是注入的 `fetch`，这个文件一个 DOM API 都不碰，Node 里就能测。
 */
import type { PageLayout } from '@uw/layout';
import type { RElement, RenderOptions } from '@uw/render-dom';
import { buildPage, serialize, textEl } from '@uw/render-dom';

/** 取一个地址的字节。失败抛错 —— 调用方逐项吞掉，一张图 / 一款字体取不到不该让整页导不出来 */
export type FetchBytes = (url: string) => Promise<{ bytes: Uint8Array; type: string }>;

/** 从宿主样式表里收来的一条 `@font-face`：族名（已去引号、小写）+ 规则原文 + 解析相对地址用的基准 */
export interface FontFaceSource {
  family: string;
  cssText: string;
  baseUrl: string;
}

/** 一串 CSS `font-family` → 各个族名（去引号、小写）。`serif` 这种通用族也在里面，不影响匹配 */
export function familiesOf(value: string): string[] {
  return value
    .split(',')
    .map((f) => normalizeFamily(f))
    .filter((f) => f !== '');
}

export function normalizeFamily(value: string): string {
  return value
    .trim()
    .replace(/^(['"])(.*)\1$/, '$2')
    .trim()
    .toLowerCase();
}

/** 元素树里每个 `font-family` 属性提到过的族名 —— 只有它们对应的 `@font-face` 才值得内联 */
export function usedFamilies(node: RElement, out = new Set<string>()): Set<string> {
  const value = node.attrs['font-family'];
  if (value !== undefined) for (const f of familiesOf(value)) out.add(f);
  for (const c of node.children) usedFamilies(c, out);
  return out;
}

/** 元素树里所有 `<image href>`（去重）。只有非 data URI 的才需要取 */
export function externalImageHrefs(node: RElement, out = new Set<string>()): Set<string> {
  const href = node.tag === 'image' ? node.attrs.href : undefined;
  if (href !== undefined && !href.startsWith('data:')) out.add(href);
  for (const c of node.children) externalImageHrefs(c, out);
  return out;
}

function replaceHrefs(node: RElement, map: ReadonlyMap<string, string>): RElement {
  const href = node.tag === 'image' ? node.attrs.href : undefined;
  const next = href === undefined ? undefined : map.get(href);
  const children = node.children.map((c) => replaceHrefs(c, map));
  const attrs = next === undefined ? node.attrs : { ...node.attrs, href: next };
  return node.text === undefined
    ? { tag: node.tag, attrs, children }
    : { tag: node.tag, attrs, children, text: node.text };
}

/** 字节 → data URI。分块 `String.fromCharCode`：一次展开几 MB 的数组会爆调用栈 */
export function dataUri(bytes: Uint8Array, type: string): string {
  let binary = '';
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
  }
  return `data:${type || 'application/octet-stream'};base64,${btoa(binary)}`;
}

const URL_PATTERN = /url\(\s*(['"]?)(.*?)\1\s*\)/g;

/** 响应没带类型（file:// 或配置不全的静态服务器）时按扩展名补，缺了类型部分浏览器不认字体 */
function fontType(url: string, type: string): string {
  if (type !== '' && type !== 'application/octet-stream') return type;
  const ext = /\.(woff2|woff|ttf|otf)(?:[?#]|$)/i.exec(url)?.[1]?.toLowerCase();
  return ext === undefined ? 'application/octet-stream' : `font/${ext}`;
}

/**
 * 一条 `@font-face` 里的 `url(…)` 全换成 data URI。有一个换成功就留下这条规则 ——
 * `src` 本来就是一串退路，取不到的那几项留着原地址也无害（图片沙箱里只是加载失败）；
 * 一个都没换成就丢掉整条，免得白白占着族名挡住本机同名字体。
 */
async function inlineFontFace(face: FontFaceSource, fetchBytes: FetchBytes): Promise<string | undefined> {
  const urls = [...face.cssText.matchAll(URL_PATTERN)].map((m) => m[2] ?? '');
  const inlined = new Map<string, string>();
  for (const raw of urls) {
    if (raw === '' || raw.startsWith('data:') || inlined.has(raw)) continue;
    try {
      const url = new URL(raw, face.baseUrl).href;
      const { bytes, type } = await fetchBytes(url);
      inlined.set(raw, dataUri(bytes, fontType(url, type)));
    } catch {
      // 跨域没开 CORS、404：这一项退回原地址
    }
  }
  const alreadyInline = urls.some((u) => u.startsWith('data:'));
  if (inlined.size === 0 && !alreadyInline) return undefined;
  return face.cssText.replace(URL_PATTERN, (whole, _q: string, raw: string) => {
    const uri = inlined.get(raw);
    return uri === undefined ? whole : `url("${uri}")`;
  });
}

export interface SnapshotOptions extends RenderOptions {
  /** 宿主样式表里的 `@font-face`，由 DOM 那一侧收集 */
  fontFaces?: readonly FontFaceSource[];
  fetchBytes?: FetchBytes;
}

/**
 * 一页 → 可以直接喂给 `<img>` 的独立 SVG 文本。`zoom` 就是导出倍率：
 * `<svg>` 的 width / height 按它出 px，viewBox 仍是 pt，与屏幕那一份同一条缩放路径。
 */
export async function pageSnapshotMarkup(page: PageLayout, opts: SnapshotOptions = {}): Promise<string> {
  let tree = buildPage(page, opts);
  const fetchBytes = opts.fetchBytes;
  if (fetchBytes !== undefined) {
    const map = new Map<string, string>();
    for (const href of externalImageHrefs(tree)) {
      try {
        const { bytes, type } = await fetchBytes(href);
        map.set(href, dataUri(bytes, type));
      } catch {
        // 取不到就留着原地址：这张图画不出来，不挡其余的
      }
    }
    if (map.size > 0) tree = replaceHrefs(tree, map);
  }
  const used = usedFamilies(tree);
  const css: string[] = [];
  for (const face of opts.fontFaces ?? []) {
    if (!used.has(face.family)) continue;
    const rule = fetchBytes === undefined ? face.cssText : await inlineFontFace(face, fetchBytes);
    if (rule !== undefined) css.push(rule);
  }
  if (css.length > 0) tree = { ...tree, children: [textEl('style', {}, css.join('\n')), ...tree.children] };
  return serialize(tree);
}
