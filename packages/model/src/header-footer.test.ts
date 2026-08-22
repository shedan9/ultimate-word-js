/**
 * 页眉页脚部件的解析：`spike-header-03.docx` 里有三份页眉三份页脚，页脚里还有真的 `{ PAGE }` 域。
 *
 * 这里只管 model 这一侧「解析出来对不对」；「哪一页画哪一份、摆在纸的什么位置」是
 * `@uw/layout` 的 `header-fixture.test.ts` 的事。
 */
import { readFileSync } from 'node:fs';
import { createDiagnosticSink } from '@uw/core';
import { OpcPackage } from '@uw/ooxml';
import { beforeAll, describe, expect, it } from 'vitest';
import type { LoadedDocument } from './load.ts';
import { loadDocument } from './load.ts';
import { paragraphText, walkBlocks } from './nodes.ts';

const FIXTURE = new URL('../../../apps/fidelity/fixtures/spike-header-03.docx', import.meta.url);

let doc: LoadedDocument;
const sink = createDiagnosticSink();

beforeAll(() => {
  doc = loadDocument(OpcPackage.open(new Uint8Array(readFileSync(FIXTURE))), sink);
});

/** 一份页眉页脚里所有段落的文字 */
function textOf(relId: string): string[] {
  const content = doc.headerFooters[relId];
  if (content === undefined) throw new Error(`没有 ${relId} 这份内容`);
  return [...walkBlocks(content.blocks)].filter((b) => b.kind === 'paragraph').map((b) => paragraphText(b));
}

describe('页眉页脚部件', () => {
  it('本节引用到的六份全部解析出来了，一份不多一份不少', () => {
    const props = doc.body.sections[0]?.props;
    const ids = [...(props?.headers ?? []), ...(props?.footers ?? [])].map((r) => r.relId);
    expect(ids).toHaveLength(6);
    expect(Object.keys(doc.headerFooters).sort()).toEqual([...ids].sort());
  });

  it('三份页眉的文字各不相同 —— 引用的 w:type 与内容对得上', () => {
    const props = doc.body.sections[0]?.props;
    const byType = Object.fromEntries((props?.headers ?? []).map((r) => [r.type, textOf(r.relId)]));
    expect(byType.first).toEqual(['首页页眉']);
    expect(byType.default).toEqual(['奇数页眉']);
    expect(byType.even).toEqual(['偶数页眉']);
  });

  it('节点 id 带部件前缀 —— 与正文撞车会让页脚的页码画到正文里去', () => {
    const [relId] = Object.keys(doc.headerFooters);
    const content = doc.headerFooters[relId as string];
    const first = content?.blocks[0];
    expect(first?.id.startsWith(`${relId}:`)).toBe(true);
    // 正文那棵树的 id 一个前缀都不带，两边因此永远不会相等
    expect(doc.body.sections[0]?.blocks[0]?.id.startsWith('p')).toBe(true);
  });

  it('级联跑过了：resolved 与 blocks 一一对应', () => {
    for (const content of Object.values(doc.headerFooters)) {
      expect(content.resolved).toHaveLength(content.blocks.length);
    }
  });

  it('页脚里的 PAGE 域也扫进了同一份 fields —— 求值那边不关心它来自哪个部件', () => {
    const pages = doc.fields.filter((f) => f.instruction.type === 'PAGE');
    expect(pages).toHaveLength(3);
    // 结果区落在页脚部件的 run 上，所以 id 带着那份页脚的关系前缀
    for (const f of pages) expect(f.resultRuns[0]).toMatch(/^rId\d+:/u);
  });

  it('一条诊断都没有', () => {
    expect(sink.list()).toHaveLength(0);
  });
});
