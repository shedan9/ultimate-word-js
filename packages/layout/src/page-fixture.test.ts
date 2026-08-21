/**
 * 分页的真值回归：`spike-page-01/02` 两份样本共 50 页，**逐页逐行**与 Word 对。
 *
 * 与 `fixture.test.ts`（真实公文，一页）分开放，因为这两份问的是不同的问题：
 * 那一份测的是**横向**（断行点）与一页之内的 y，这一份测的是**分页规则**本身 ——
 * 孤行寡行在第几行断、keepNext 把哪一行拽走、页首的段前间距算不算。
 *
 * 样本的版心刻意做成「一页恰好 11 行、一行 18 个汉字、固定行距 20pt」，于是行高与字宽
 * 都不依赖任何待标定的度量，阶梯靠垫行的条数移动断页点。规则本身是怎么标定出来的
 * （3 × 2 × 3 种组合逐页比对）见 `apps/fidelity` 的 `spike:page`；这里只当**闸门**：
 * 50 页全对，少一页都算退步。
 *
 * 跨平台：docx 与 truth.json 都入库，度量走随库的度量包，所以 CI 上也跑得了。
 */
import { readFileSync } from 'node:fs';
import { createDiagnosticSink } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { DocumentLayout } from './page.ts';
import { layoutDocument } from './page.ts';

const FIXTURES = new URL('../../../apps/fidelity/fixtures/', import.meta.url);
/** L3 的判据 */
const TOLERANCE_PT = 0.5;

interface Truth {
  pageCount: number;
  pages: { lines: { y: number; text: string }[] }[];
}

interface Line {
  y: number;
  text: string;
}

function layout(name: string): { doc: DocumentLayout; truth: Truth; diagnostics: number } {
  const sink = createDiagnosticSink();
  const bytes = new Uint8Array(readFileSync(new URL(`${name}.docx`, FIXTURES)));
  const loaded = loadDocument(OpcPackage.open(bytes), sink);
  const registry = new FontRegistry();
  for (const pack of loadBundledPacks()) registry.registerMetrics(pack);
  const measurer = createTextMeasurer(registry, {
    candidates: (family) => fontNameCandidates(loaded.fonts, family),
    diagnostics: sink,
  });
  return {
    doc: layoutDocument(loaded.resolved, { measurer, settings: loaded.cascade.settings }),
    truth: JSON.parse(readFileSync(new URL(`${name}.truth.json`, FIXTURES), 'utf8')) as Truth,
    diagnostics: sink.list().length,
  };
}

/** 一页上的行：文字 + 基线的绝对 y（pt）。空段落在 PDF 里不落墨，真值没有它，这边也要滤掉 */
function linesOf(doc: DocumentLayout, page: number): Line[] {
  const p = doc.pages[page];
  if (p === undefined) return [];
  return p.blocks
    .filter((b) => b.kind === 'paragraph')
    .flatMap((b) => b.lines)
    .map((placed) => ({
      y: (p.geometry.content.y + placed.y + placed.line.baseline) / 20,
      text: placed.line.fragments
        .map((f) => f.text)
        .join('')
        .replace(/\s+$/u, ''),
    }))
    .filter((l) => l.text !== '');
}

describe.each([
  ['spike-page-01', '孤行寡行的下限'],
  ['spike-page-02', 'keepNext / keepLines / 页首段前间距'],
])('%s · %s', (name) => {
  const { doc, truth, diagnostics } = layout(name);

  it('L0：页数与 Word 一致', () => {
    expect(doc.pages).toHaveLength(truth.pageCount);
  });

  it('每一页的每一行都与 Word 逐字一致 —— 断页点错一行，后面每一页都会跟着错', () => {
    const ours = doc.pages.map((_, i) => linesOf(doc, i).map((l) => l.text));
    const theirs = truth.pages.map((p) => p.lines.map((l) => l.text));
    expect(ours).toEqual(theirs);
  });

  it(`L3：每一行的基线 y 与真值差 < ${TOLERANCE_PT}pt`, () => {
    const worst = { page: -1, line: -1, delta: 0, text: '' };
    doc.pages.forEach((_, i) => {
      linesOf(doc, i).forEach((l, k) => {
        const delta = Math.abs(l.y - (truth.pages[i]?.lines[k]?.y ?? l.y));
        if (delta > worst.delta) Object.assign(worst, { page: i, line: k, delta, text: l.text });
      });
    });
    expect(
      worst.delta,
      `最坏的是第 ${worst.page + 1} 页第 ${worst.line + 1} 行「${worst.text}」`,
    ).toBeLessThan(TOLERANCE_PT);
  });

  it('一条诊断都没有', () => {
    expect(diagnostics).toBe(0);
  });
});
