/**
 * 表格**拆行**的真值回归：`spike-table-04`，14 页逐行与 Word 对。
 *
 * 在这份样本之前，拆行的四条规则（切在哪一页、边距怎么分、`w:trHeight` 的富余归谁、
 * 头片认不认 `w:vAlign`）全是「哪种最省地方」猜的，**三条猜反了**。规则本身与证据表见
 * `table-split.ts` 的 `TABLE_SPLIT_RULES`，怎么排组合标定出来的见 `apps/fidelity` 的
 * `spike:table-split`（16 种组合逐页跑，唯一满分）。
 *
 * 两处判法与别的 fixture 测试不同，都是被样本逼出来的：
 *
 * ① **按 y 分桶重做分行**：真值的一「行」是 `extract-truth.ts` 按 y 分桶拼出来的，
 *    它不认识段落也不认识表格 —— 一行两个单元格的文字会拼成同一行。所以这边也得
 *    把整页摊平再按同样的规则分桶，否则比的根本不是同一样东西；
 * ② **比逐行增量而不是绝对 y**（与 `image-fixture.test.ts` 同一个理由）：Word 自己的
 *    行距带着 ±0.12pt 抖动（同一份文档里 15.48 / 15.60 / 15.63 都出现过），十八行叠起来
 *    累到 0.58pt，早已越过 L3 判据，而那与拆行毫无关系。每页**第一行**仍按绝对 y 比 ——
 *    「这一片从哪儿起排」正是这份样本要量的东西。
 *
 * 跨平台：docx 与 truth.json 都入库，度量走随库的度量包，CI 上照跑。
 */
import { readFileSync } from 'node:fs';
import { createDiagnosticSink } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { layoutDocumentWithFields } from './fields.ts';
import type { PageLayout } from './page.ts';
import type { Chunk } from './test-flatten.ts';
import { chunksOf } from './test-flatten.ts';

const FIXTURES = new URL('../../../apps/fidelity/fixtures/', import.meta.url);
/** L3 的判据 */
const TOLERANCE_PT = 0.5;
/** 与 `extract-truth.ts` 的 `groupLines()` 同一个容差（pt） */
const LINE_TOL = 1.2;

interface Truth {
  pageCount: number;
  pages: { lines: { y: number; x: number; text: string }[] }[];
}

interface Line {
  y: number;
  text: string;
}

const strip = (s: string): string => s.replace(/\s+/gu, '');

/** 按 y 分桶 → 桶内按 x 排 → 拼成一行 */
function groupLines(pieces: readonly Chunk[]): Line[] {
  const sorted = [...pieces].sort((a, b) => a.y - b.y || a.x - b.x);
  const lines: Line[] = [];
  let cur: Chunk[] = [];
  let curY = Number.NaN;

  const flush = (): void => {
    if (cur.length === 0) return;
    const text = strip(
      [...cur]
        .sort((a, b) => a.x - b.x)
        .map((p) => p.text)
        .join(''),
    );
    if (text !== '') lines.push({ y: curY, text });
    cur = [];
  };

  for (const p of sorted) {
    if (cur.length === 0) curY = p.y;
    else if (Math.abs(p.y - curY) > LINE_TOL) {
      flush();
      curY = p.y;
    }
    cur.push(p);
  }
  flush();
  return lines;
}

const bytes = new Uint8Array(readFileSync(new URL('spike-table-04.docx', FIXTURES)));
const truth = JSON.parse(readFileSync(new URL('spike-table-04.truth.json', FIXTURES), 'utf8')) as Truth;

const sink = createDiagnosticSink();
const loaded = loadDocument(OpcPackage.open(bytes), sink);
const registry = new FontRegistry();
for (const pack of loadBundledPacks()) registry.registerMetrics(pack);
const doc = layoutDocumentWithFields(loaded.resolved, loaded.fields, {
  measurer: createTextMeasurer(registry, {
    candidates: (family) => fontNameCandidates(loaded.fonts, family),
    diagnostics: sink,
  }),
  settings: loaded.cascade.settings,
  headerFooters: loaded.headerFooters,
}).layout;

describe('spike-table-04：拆行逐页比真值', () => {
  it('页数与 Word 一致', () => {
    expect(doc.pages).toHaveLength(truth.pages.length);
  });

  it('每一页的行文字与 Word 一致', () => {
    for (const [i, page] of doc.pages.entries()) {
      const ours = groupLines(chunksOf(page as PageLayout)).map((l) => l.text);
      const theirs = (truth.pages[i]?.lines ?? []).map((l) => strip(l.text));
      expect(ours, `第 ${i + 1} 页`).toEqual(theirs);
    }
  });

  it('每一页的首行 y 与逐行增量都在判据内', () => {
    const bad: string[] = [];
    for (const [i, page] of doc.pages.entries()) {
      const ours = groupLines(chunksOf(page as PageLayout));
      const theirs = truth.pages[i]?.lines ?? [];
      if (ours.length !== theirs.length) continue; // 上一条已经报过了
      const first = Math.abs((ours[0]?.y ?? 0) - (theirs[0]?.y ?? 0));
      if (first > TOLERANCE_PT) bad.push(`第 ${i + 1} 页首行差 ${first.toFixed(3)}pt`);
      for (let k = 1; k < ours.length; k++) {
        const a = (ours[k]?.y ?? 0) - (ours[k - 1]?.y ?? 0);
        const b = (theirs[k]?.y ?? 0) - (theirs[k - 1]?.y ?? 0);
        if (Math.abs(a - b) > TOLERANCE_PT)
          bad.push(`第 ${i + 1} 页第 ${k + 1} 行的行距差 ${Math.abs(a - b).toFixed(3)}pt`);
      }
    }
    expect(bad).toEqual([]);
  });
});
