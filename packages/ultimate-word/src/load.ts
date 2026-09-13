/**
 * `UltimateWord.load()` —— 唯一的异步入口（api.md §3）。
 *
 * 异步只因为**拿字节**可能是异步的（Blob / Response / URL）；拿到字节之后的
 * 解包 → 级联 → 度量 → 断行 → 分页 → 域求值全是同步的，与 `apps/playground` 里那条链
 * 一字不差。这一层不做任何排版决定，它只负责把六个包接成一句话。
 *
 * 结构性错误（不是 zip、缺 document.xml）从 `OpcPackage.open` / `loadDocument` 里**抛**出来，
 * 内容问题（缺字体、不认识的元素）记进 `doc.diagnostics` 并照常排版 —— 架构原则 1.5，
 * 这里不拦也不转译。
 */
import { createDiagnosticSink } from '@uw/core';
import type { FontRegistry } from '@uw/fonts';
import { createTextMeasurer } from '@uw/fonts';
import { layoutDocumentWithFields } from '@uw/layout';
import {
  DEFAULT_SECTION_PROPS,
  fieldHyperlinks,
  fontNameCandidates,
  loadDocument,
  resolveBody,
  scanFields,
} from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import { UwDocument } from './document.ts';

/** 字符串是 URL，走 `fetch`。`File` 是 `Blob` 的子类，不必单列 */
export type LoadSource = ArrayBuffer | Uint8Array | Blob | Response | string;

export interface LoadOptions {
  /** 按文档覆盖字体注册表，默认用全局那份（`UltimateWord.fonts.registry`） */
  fonts?: FontRegistry;
  /** 只作用于取字节那一步（fetch / Blob 读取）；排版是同步的，中途停不下来 */
  signal?: AbortSignal;
}

async function toBytes(source: LoadSource, signal: AbortSignal | undefined): Promise<Uint8Array> {
  if (source instanceof Uint8Array) return source;
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  if (typeof source === 'string') {
    const init: RequestInit = signal === undefined ? {} : { signal };
    const res = await fetch(source, init);
    if (!res.ok) throw new Error(`取不到 ${source}：HTTP ${res.status}`);
    return new Uint8Array(await res.arrayBuffer());
  }
  // Blob 与 Response 都有 arrayBuffer()；Response 要先看状态，一个 404 页面解不出 zip
  if (typeof Response !== 'undefined' && source instanceof Response) {
    if (!source.ok) throw new Error(`响应不成功：HTTP ${source.status}`);
    return new Uint8Array(await source.arrayBuffer());
  }
  signal?.throwIfAborted();
  return new Uint8Array(await source.arrayBuffer());
}

export async function load(
  source: LoadSource,
  registry: FontRegistry,
  options: LoadOptions = {},
): Promise<UwDocument> {
  const bytes = await toBytes(source, options.signal);
  options.signal?.throwIfAborted();

  const sink = createDiagnosticSink();
  const loaded = loadDocument(OpcPackage.open(bytes), sink);
  const measurer = createTextMeasurer(options.fonts ?? registry, {
    // 「黑体」→「SimHei」的桥在文档自己的 fontTable 里，fonts 不认识 model，所以在这儿接
    candidates: (family) => fontNameCandidates(loaded.fonts, family),
    diagnostics: sink,
  });
  const result = layoutDocumentWithFields(loaded.resolved, loaded.fields, {
    measurer,
    settings: loaded.cascade.settings,
    headerFooters: loaded.headerFooters,
    diagnostics: sink,
  });
  return new UwDocument({
    loaded,
    reflow(body) {
      const fields = scanFields(body, sink);
      for (const hf of Object.values(loaded.headerFooters))
        fields.push(
          ...scanFields(
            {
              sections: [
                {
                  id: hf.relId,
                  props: loaded.body.sections[0]?.props ?? DEFAULT_SECTION_PROPS,
                  blocks: hf.blocks,
                },
              ],
            },
            sink,
          ),
        );
      const resolved = resolveBody(loaded.cascade, body, { hyperlinks: fieldHyperlinks(fields) });
      const result = layoutDocumentWithFields(resolved, fields, {
        measurer,
        settings: loaded.cascade.settings,
        headerFooters: loaded.headerFooters,
        diagnostics: sink,
      });
      return { loaded: { ...loaded, body, resolved, fields }, layout: result.layout, values: result.values };
    },
    layout: result.layout,
    fieldValues: result.values,
    diagnostics: sink.list(),
  });
}
