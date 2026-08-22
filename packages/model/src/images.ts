/**
 * 图片资源 —— 把 `ImageRef.id` 换成一份真的字节。
 *
 * ## 为什么单独摊平成一张表，而不是把字节挂在节点上
 *
 * 与页眉页脚同一个道理（load.ts）：同一张图常被引用很多次（每页页眉里的 logo、
 * 公文里反复出现的分隔线），挂在节点上要么复制几份字节，要么变成跨节点引用 ——
 * 后者正是原则 1.1 挡的反向指针。摊平之后节点上只留一个字符串 key，
 * 整棵树仍然可结构化克隆，几 MB 的字节也只存一份。
 *
 * ## 为什么按「引用」收，而不是把 `RelType.IMAGE` 全捞出来
 *
 * 与页眉页脚同一条经验：Word 删掉一张图之后，`media/image3.png` 常常还留在包里。
 * 按类型全捞会把这些废弃的字节全读进内存（图片是包里最大的东西，一张扫描件几 MB），
 * 而它们一个像素都不会画出来。所以这里先走一遍节点树，只收真被引用到的。
 *
 * ## 外链图片（`r:link` / `TargetMode="External"`）只给地址
 *
 * 字节不在包里，取它要发网络请求 —— 那是宿主该决定的事（离线预览、内网隔离、
 * 防跟踪都可能要求不发），模型层擅自发请求会把一个纯计算的库变成网络客户端。
 * 所以这里只把 URL 带出来，画不画由渲染层的 `imageHref` 决定。
 */
import type { DiagnosticSink } from '@uw/core';
import type { OpcPart, Relationships } from '@uw/ooxml';
import type { BlockNode, ImageRef, PropSet } from './nodes.ts';
import { walkBlocks } from './nodes.ts';

/** 一张图的字节与身份。`bytes` 与 `url` **至少有一个**，两个都缺的不会进表 */
export interface ImageResource {
  /** 与 `ImageRef.id` 同一个 key（部件前缀 + 关系 id） */
  id: string;
  relId: string;
  /** 部件名，如 `/word/media/image1.png`。外链图片没有 */
  part?: string;
  /**
   * 内容类型（`image/png`）。查不到时按扩展名兜底 —— 少数生成器不写
   * `[Content_Types].xml` 里的默认扩展名项，而 `<image>` 的 data URI 少了它就画不出来。
   */
  contentType: string;
  /** 外链图片的地址 */
  url?: string;
  /** 字节。**外链图片没有**，见文件头 */
  bytes?: Uint8Array;
}

/** 收图片只需要包的这么点能力。声明成结构类型是为了单测能拿个字面量当包用 */
export interface ImagePartSource {
  rels(partName?: string): Relationships;
  part(name: string): OpcPart | undefined;
}

/** 一个部件里的内容 + 它的 id 前缀。前缀的由来见 parse-body.ts 的 `Ctx.idPrefix` */
export interface ImageScope<S extends PropSet> {
  part: string;
  idPrefix: string;
  blocks: readonly BlockNode<S>[];
}

/** 扩展名 → 内容类型的兜底表。只列 `<img>` 画得出来的那些，其余走 `application/octet-stream` */
const BY_EXTENSION: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
};

/**
 * 走一遍各部件的节点树，把被引用到的图片读出来。
 *
 * 同一个 key 只读一次（`out[id]`）—— 一张 logo 在页眉里出现 200 次也只有一份字节。
 */
export function collectImages<S extends PropSet>(
  pkg: ImagePartSource,
  scopes: readonly ImageScope<S>[],
  diagnostics: DiagnosticSink,
): Record<string, ImageResource> {
  const out: Record<string, ImageResource> = {};
  for (const scope of scopes) {
    for (const ref of imageRefs(scope.blocks)) {
      if (out[ref.id] !== undefined) continue;
      const res = resolveImage(pkg, scope.part, ref, diagnostics);
      if (res !== undefined) out[ref.id] = res;
    }
  }
  return out;
}

/** 一列块里出现的全部图片引用，按文档顺序 */
export function* imageRefs<S extends PropSet>(blocks: readonly BlockNode<S>[]): Generator<ImageRef> {
  for (const block of walkBlocks(blocks)) {
    if (block.kind !== 'paragraph') continue;
    for (const run of block.runs) {
      for (const c of run.content) {
        if (c.kind === 'object' && c.image !== undefined) yield c.image;
      }
    }
  }
}

function resolveImage(
  pkg: ImagePartSource,
  part: string,
  ref: ImageRef,
  diagnostics: DiagnosticSink,
): ImageResource | undefined {
  const rel = pkg.rels(part).byId(ref.relId);
  if (rel === undefined) {
    // 悬空引用不是结构性错误：Word 自己也留过这种。当成「这张图没有」继续画（原则 1.5）
    diagnostics.warn('image-missing', `图片引用 ${ref.relId} 在 ${part} 的关系表里查不到`, {
      part,
    });
    return undefined;
  }
  if (rel.targetMode === 'External' || rel.target === undefined) {
    return { id: ref.id, relId: ref.relId, contentType: guessType(rel.rawTarget), url: rel.rawTarget };
  }

  const opcPart = pkg.part(rel.target);
  if (opcPart === undefined) {
    diagnostics.warn('image-missing', `图片 ${rel.target} 在关系表里有，包里却没有这个部件`, {
      part: rel.target,
    });
    return undefined;
  }
  return {
    id: ref.id,
    relId: ref.relId,
    part: rel.target,
    contentType: opcPart.contentType ?? guessType(rel.target),
    bytes: opcPart.bytes,
  };
}

function guessType(name: string): string {
  const dot = name.lastIndexOf('.');
  const ext = dot < 0 ? '' : name.slice(dot + 1).toLowerCase();
  return BY_EXTENSION[ext] ?? 'application/octet-stream';
}
