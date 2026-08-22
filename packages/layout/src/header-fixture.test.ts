/**
 * 页眉页脚的真值回归：`spike-header-01/02/03` 三份样本共 12 页，**逐页逐行**与 Word 对。
 *
 * 比对方式与 `page-fixture.test.ts` 有一处关键差别：这里把页眉、正文、页脚的行**混在一起
 * 按 y 排序**再比。真值来自 PDF，而 PDF 里没有「这是页眉」这回事 —— 页眉排错了位置，
 * 表现就是这张按 y 排好的表对不上，而不是某个单独的指标。
 *
 * 三份样本各自负责一格（几何规则怎么标定出来的见 `apps/fidelity` 的 `spike:header`）：
 * 01 页眉页脚各一行（放得下，正文不该动）、02 各三行（放不下，该顶开正文并少排一行）、
 * 03 开「首页不同 + 奇偶页不同」且页脚里是**真的** `{ PAGE }` 域。
 * 03 因此同时是「域求值 → 页眉」这条链的端到端闸门：页码算错会直接变成文字对不上。
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
import { layoutDocumentWithFields } from './fields.ts';
import type { DocumentLayout, PageLayout, PlacedBlock } from './page.ts';

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
    doc: layoutDocumentWithFields(loaded.resolved, loaded.fields, {
      measurer,
      settings: loaded.cascade.settings,
      headerFooters: loaded.headerFooters,
    }).layout,
    truth: JSON.parse(readFileSync(new URL(`${name}.truth.json`, FIXTURES), 'utf8')) as Truth,
    diagnostics: sink.list().length,
  };
}

/** 一个容器里每一行的绝对基线（pt）与文字。空段落在 PDF 里不落墨，真值没有它，这边也要滤掉 */
function linesOf(blocks: readonly PlacedBlock[], y0: number): Line[] {
  const out: Line[] = [];
  for (const b of blocks) {
    if (b.kind !== 'paragraph') continue;
    for (const placed of b.lines) {
      const text = placed.line.fragments
        .map((f) => f.text)
        .join('')
        .replace(/\s+$/u, '');
      if (text !== '') out.push({ y: (y0 + placed.y + placed.line.baseline) / 20, text });
    }
  }
  return out;
}

function pageLines(page: PageLayout): Line[] {
  return [
    ...(page.header === undefined ? [] : linesOf(page.header.blocks, page.header.y)),
    ...linesOf(page.blocks, page.geometry.content.y),
    ...(page.footer === undefined ? [] : linesOf(page.footer.blocks, page.footer.y)),
  ].sort((a, b) => a.y - b.y);
}

describe.each([
  ['spike-header-01', '矮页眉页脚：放得下就不动正文'],
  ['spike-header-02', '高页眉页脚：顶开版心，一页少排三行'],
  ['spike-header-03', '首页 / 奇偶页各用哪一份 + 页脚里的 { PAGE }'],
])('%s · %s', (name) => {
  const { doc, truth, diagnostics } = layout(name);

  it('L0：页数与 Word 一致', () => {
    expect(doc.pages).toHaveLength(truth.pageCount);
  });

  it('每一页的每一行都与 Word 逐字一致（页眉 / 正文 / 页脚混在一起按 y 排）', () => {
    const ours = doc.pages.map((p) => pageLines(p).map((l) => l.text));
    const theirs = truth.pages.map((p) => p.lines.map((l) => l.text));
    expect(ours).toEqual(theirs);
  });

  it(`L3：每一行的基线 y 与真值差 < ${TOLERANCE_PT}pt`, () => {
    const worst = { page: -1, line: -1, delta: 0, text: '' };
    doc.pages.forEach((p, i) => {
      pageLines(p).forEach((l, k) => {
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

describe('spike-header-03 · 选择规则落到具体的页上', () => {
  const { doc } = layout('spike-header-03');

  it('首页用 first，之后按显示页码的奇偶交替 —— 三份页眉指的是三个不同的部件', () => {
    const ids = doc.pages.map((p) => p.header?.relId);
    expect(new Set(ids).size).toBe(3);
    // 首页单独一份；第 3、5 页共用奇数那一份；第 2、4 页共用偶数那一份
    expect(ids[0]).not.toBe(ids[1]);
    expect(ids[2]).toBe(ids[4]);
    expect(ids[1]).toBe(ids[3]);
    expect(ids[0]).not.toBe(ids[2]);
  });

  it('页脚里的 PAGE 每页各不相同，且带 field 标记（要能复制、能被 Ctrl+F 搜到）', () => {
    const texts = doc.pages.map((p) => {
      const block = p.footer?.blocks[0];
      if (block?.kind !== 'paragraph') throw new Error('页脚的第一块不是段落');
      return block.lines[0]?.line.fragments.map((f) => f.text).join('');
    });
    expect(texts).toEqual(['首页第1', '偶数第2', '奇数第3', '偶数第4', '奇数第5']);

    const block = doc.pages[2]?.footer?.blocks[0];
    if (block?.kind !== 'paragraph') throw new Error('页脚的第一块不是段落');
    expect(block.lines[0]?.line.fragments.at(-1)?.field).toBe(true);
  });

  it('版心没有被矮页眉挤过 —— 顶仍是 w:top（1134 twips = 20mm）', () => {
    expect(doc.pages[0]?.geometry.content.y).toBe(1134);
  });
});

describe('spike-header-02 · 高页眉页脚把版心挤窄', () => {
  const { doc } = layout('spike-header-02');

  it('版心顶 = 页眉底（页眉距 + 三行），不再是 w:top', () => {
    const page = doc.pages[0];
    // 页眉距 10mm = 567 twips，三行固定行距 20pt = 1200 twips
    expect(page?.header?.y).toBe(567);
    expect(page?.header?.height).toBe(1200);
    expect(page?.geometry.content.y).toBe(567 + 1200);
  });

  it('版心底 = 页脚顶（纸底 − 页脚距 − 三行）', () => {
    const page = doc.pages[0];
    const paper = page?.geometry.height ?? 0;
    expect(page?.footer?.y).toBe(paper - 567 - 1200);
    expect((page?.geometry.content.y ?? 0) + (page?.geometry.content.height ?? 0)).toBe(paper - 567 - 1200);
  });
});
