/**
 * 图片字节 → `<image href="…">` 能用的地址。
 *
 * ## 为什么这一层不认识 `@uw/model`
 *
 * 入参是**结构类型**（`{ contentType, bytes?, url? }`），`ImageResource` 正好对得上但
 * 这里不 import 它 —— 渲染层的依赖只到 `@uw/layout` 为止（架构 §3.1 的单向依赖）。
 * 换个来源（数据库里的图、宿主自己解好的 blob）也能直接喂进来。
 *
 * ## 为什么默认是 data URI 而不是 blob URL
 *
 * data URI 在**任何地方**都成立：Node 里落盘成 HTML（`apps/fidelity` 的 preview）、
 * 快照测试、将来 Worker 里预生成标记文本，都不需要 `URL.createObjectURL` 这种
 * 浏览器 API。代价是 base64 撑大 1/3，而且没有释放这一说 —— 浏览器里长期挂着几十张
 * 大图时，宿主该自己传一个 blob URL 版本的 `imageHref`（`@uw/render-dom/dom` 那一侧）。
 *
 * ## 画不出来的类型宁可不画
 *
 * EMF / WMF 是 Windows 图元文件，`<img>` 与 SVG 的 `<image>` 都不认 —— 而 Word 文档里
 * 它们很常见（从 Office 剪贴板粘过来的图默认就是 EMF）。给它一个画不出来的 href
 * 会得到一个「破图」图标；返回 undefined 则走占位框那条路，**框的尺寸是对的**，
 * 页面不会因此错位。
 */

/** 一张图的最小描述。`@uw/model` 的 `ImageResource` 结构上就是它 */
export interface ImageBytes {
  contentType: string;
  bytes?: Uint8Array;
  /** 外链图片（`r:link`）的地址 */
  url?: string;
}

/**
 * 浏览器画得出来的位图类型。
 *
 * 这张表是**渲染层的事实**（`<image>` 支持什么），不是文档格式的事实 ——
 * 所以它在这里，不在 `@uw/model`。
 */
const DISPLAYABLE = new Set([
  'image/png',
  'image/jpeg',
  'image/gif',
  'image/bmp',
  'image/webp',
  'image/svg+xml',
]);

export function isDisplayable(contentType: string): boolean {
  return DISPLAYABLE.has(contentType);
}

/**
 * 一张图 → href。画不出来的类型、以及既没有字节也没有地址的，返回 undefined。
 *
 * 外链图片**原样给 URL**：要不要真的去取它（离线预览、内网隔离、防跟踪）是宿主的决定，
 * 这一层只负责把地址递出去。
 */
export function imageHrefOf(image: ImageBytes | undefined): string | undefined {
  if (image === undefined) return undefined;
  if (image.bytes !== undefined && isDisplayable(image.contentType)) {
    return `data:${image.contentType};base64,${base64(image.bytes)}`;
  }
  return image.url;
}

/**
 * 一张「图片表」→ `RenderOptions.imageHref`。
 *
 * 直接把 `LoadedDocument.images` 传进来即可。**每张图只编码一次**：同一张 logo 在
 * 两百页的页眉里各出现一次，base64 只算一遍（几 MB 的图编码一次是毫秒级，两百次不是）。
 */
export function imageHrefResolver(
  images: Readonly<Record<string, ImageBytes>>,
): (id: string) => string | undefined {
  const cache = new Map<string, string | undefined>();
  return (id) => {
    if (cache.has(id)) return cache.get(id);
    const href = imageHrefOf(images[id]);
    cache.set(id, href);
    return href;
  };
}

const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

/**
 * 自己写 base64，不用 `Buffer`（Node 才有）也不用 `btoa`（要先把字节拼成 latin1 字符串，
 * 几 MB 的图会先造一个几 MB 的字符串）。
 *
 * 分块拼字符串是为了不把一个几百万字符的字符串一次性接起来 —— V8 的绳索字符串扛得住，
 * 但每块 8190 字节（3 的倍数，正好 10920 个 base64 字符）在实测里明显更快也更省。
 */
function base64(bytes: Uint8Array): string {
  const parts: string[] = [];
  const CHUNK = 8190;
  for (let start = 0; start < bytes.length; start += CHUNK) {
    const end = Math.min(start + CHUNK, bytes.length);
    let out = '';
    let i = start;
    for (; i + 2 < end; i += 3) {
      const n = ((bytes[i] as number) << 16) | ((bytes[i + 1] as number) << 8) | (bytes[i + 2] as number);
      out += B64[(n >> 18) & 63] as string;
      out += B64[(n >> 12) & 63] as string;
      out += B64[(n >> 6) & 63] as string;
      out += B64[n & 63] as string;
    }
    // 尾巴：剩 1 字节补两个 =，剩 2 字节补一个
    const rest = end - i;
    if (rest === 1) {
      const n = (bytes[i] as number) << 16;
      out += `${B64[(n >> 18) & 63] as string}${B64[(n >> 12) & 63] as string}==`;
    } else if (rest === 2) {
      const n = ((bytes[i] as number) << 16) | ((bytes[i + 1] as number) << 8);
      out += `${B64[(n >> 18) & 63] as string}${B64[(n >> 12) & 63] as string}${B64[(n >> 6) & 63] as string}=`;
    }
    parts.push(out);
  }
  return parts.join('');
}
