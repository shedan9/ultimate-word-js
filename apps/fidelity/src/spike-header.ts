#!/usr/bin/env node
/**
 * 页眉页脚穿刺：框摆在纸的哪儿、它反过来把版心挤成多高。
 *
 * 与 `spike:page` 同一个路子 —— **不反推系数**，把整台引擎跑一遍再与真值逐页对。
 * 理由也一样：这里不是一个数，是两条互相纠缠的判断（页脚定位错了会让页脚整体偏一个
 * 页脚高度，版心挤不挤又会改每一页装几行），单独看任何一条都会被另一条污染。
 *
 * 三份样本（`spike-header-01/02/03`）的版心与 `spike-page-01` 同一套：120×120mm 的纸、
 * 四边 20mm、页眉页脚距各 10mm、固定行距 20pt 仿宋 12pt，于是一页恰好 11 行、行高恰好 20pt。
 * 01 的页眉页脚各一行（放得下，不该动正文），02 各三行（放不下，该顶开正文），
 * 03 开「首页不同 + 奇偶页不同」并在页脚里放**真的** `{ PAGE }` 域。
 *
 * 三条结论（判据见 `@uw/layout` 的 `HEADER_RULES`）：
 *
 * 1. 页眉框顶 = `w:header`（到**纸**顶）
 * 2. 页脚量的是框**底**：框底 = 纸高 − `w:footer`，与页眉不对称
 * 3. 页边距是**最小值**：版心顶 = max(`w:top`, 页眉底)、版心底 = min(纸高 − `w:bottom`, 页脚顶)
 *
 * 这个脚本**不需要 Word**：docx 与 truth.json 都入库了，度量走随库的度量包。
 * 重新造样本才要 Windows（`pnpm truth spike-header-01 spike-header-02 spike-header-03`）。
 *
 *   node src/spike-header.ts
 */
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDiagnosticSink, twipsToPt } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import type { DocumentLayout, HeaderRules, PageLayout, PlacedBlock } from '@uw/layout';
import { HEADER_RULES, layoutDocumentWithFields } from '@uw/layout';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import type { WordTruth } from './truth-types.ts';

const APP_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const FIXTURES = ['spike-header-01', 'spike-header-02', 'spike-header-03'];
/** L3 的判据。页眉页脚的基线与正文的一起比 */
const TOLERANCE_PT = 0.5;

const CANDIDATES = {
  footerAnchor: ['bottom', 'top'],
  squeeze: ['both', 'none', 'top', 'bottom'],
} as const satisfies { [K in keyof HeaderRules]: readonly HeaderRules[K][] };

interface Line {
  y: number;
  text: string;
}

interface Result {
  okPages: number;
  totalPages: number;
  worstY: number;
  firstBad?: { page: number; ours: string[]; theirs: string[]; deltaY?: number };
}

interface Loaded {
  name: string;
  bytes: Uint8Array;
  truth: WordTruth;
}

async function load(name: string): Promise<Loaded> {
  const bytes = new Uint8Array(await readFile(path.join(APP_ROOT, 'fixtures', `${name}.docx`)));
  const truth = JSON.parse(
    await readFile(path.join(APP_ROOT, 'fixtures', `${name}.truth.json`), 'utf8'),
  ) as WordTruth;
  return { name, bytes, truth };
}

/** 一块里每一行的绝对基线 + 文字。`y0` 是这块所在容器（版心 / 页眉框）的顶 */
function linesOf(blocks: readonly PlacedBlock[], y0: number): Line[] {
  const out: Line[] = [];
  for (const b of blocks) {
    if (b.kind !== 'paragraph') continue;
    for (const placed of b.lines) {
      const text = placed.line.fragments
        .map((f) => f.text)
        .join('')
        .replace(/\s+$/u, '');
      // 空段落在 PDF 里不落墨，真值也就没有那一行
      if (text !== '') out.push({ y: twipsToPt(y0 + placed.y + placed.line.baseline), text });
    }
  }
  return out;
}

/**
 * 一页上**所有**的行，按 y 排序 —— 页眉、正文、页脚一视同仁。
 *
 * 真值来自 PDF，PDF 里没有「这是页眉」这回事，所以比对必须在同一个坐标系里做：
 * 页眉排错位置在这里表现为「行的顺序或 y 对不上」，而不是某个单独的指标。
 */
function pageLines(page: PageLayout): Line[] {
  const out = [
    ...(page.header === undefined ? [] : linesOf(page.header.blocks, page.header.y)),
    ...linesOf(page.blocks, page.geometry.content.y),
    ...(page.footer === undefined ? [] : linesOf(page.footer.blocks, page.footer.y)),
  ];
  return out.sort((a, b) => a.y - b.y);
}

function layoutOf(fixture: Loaded, headerRules: HeaderRules): DocumentLayout {
  const sink = createDiagnosticSink();
  const doc = loadDocument(OpcPackage.open(fixture.bytes), sink);
  const registry = new FontRegistry();
  for (const pack of loadBundledPacks()) registry.registerMetrics(pack);
  const measurer = createTextMeasurer(registry, {
    candidates: (family) => fontNameCandidates(doc.fonts, family),
    diagnostics: sink,
  });
  return layoutDocumentWithFields(doc.resolved, doc.fields, {
    measurer,
    settings: doc.cascade.settings,
    headerFooters: doc.headerFooters,
    headerRules,
  }).layout;
}

function run(fixture: Loaded, headerRules: HeaderRules): Result {
  const out = layoutOf(fixture, headerRules);
  const result: Result = { okPages: 0, totalPages: fixture.truth.pages.length, worstY: 0 };

  out.pages.forEach((page, i) => {
    const ours = pageLines(page);
    const theirs = fixture.truth.pages[i]?.lines ?? [];
    const sameText = ours.length === theirs.length && ours.every((l, k) => l.text === theirs[k]?.text);
    const worstY = sameText ? Math.max(0, ...ours.map((l, k) => Math.abs(l.y - (theirs[k]?.y ?? l.y)))) : 0;
    if (!sameText || worstY > TOLERANCE_PT) {
      result.firstBad ??= {
        page: i,
        ours: ours.map((l) => l.text),
        theirs: theirs.map((l) => l.text),
        ...(sameText ? { deltaY: worstY } : {}),
      };
      return;
    }
    result.okPages += 1;
    if (worstY > result.worstY) result.worstY = worstY;
  });
  return result;
}

const fixtures = await Promise.all(FIXTURES.map(load));

const combos: HeaderRules[] = [];
for (const footerAnchor of CANDIDATES.footerAnchor) {
  for (const squeeze of CANDIDATES.squeeze) combos.push({ footerAnchor, squeeze });
}

const label = (r: HeaderRules): string => `页脚量 ${r.footerAnchor.padEnd(6)} · 挤版心 ${r.squeeze}`;
const isDefault = (r: HeaderRules): boolean =>
  r.footerAnchor === HEADER_RULES.footerAnchor && r.squeeze === HEADER_RULES.squeeze;

console.log(`\n页眉页脚几何 · ${combos.length} 种组合 × ${fixtures.length} 份样本，逐页比对\n`);
const LABEL_W = 34;
console.log(`${'组合'.padEnd(LABEL_W)}${fixtures.map((f) => f.name.padStart(18)).join('')}   总对上`);
console.log('-'.repeat(LABEL_W + 18 * fixtures.length + 8));

const scored: { rules: HeaderRules; ok: number; total: number; results: Result[] }[] = [];
for (const rules of combos) {
  const results = fixtures.map((f) => run(f, rules));
  const ok = results.reduce((n, r) => n + r.okPages, 0);
  const total = results.reduce((n, r) => n + r.totalPages, 0);
  scored.push({ rules, ok, total, results });
  console.log(
    (isDefault(rules) ? '→ ' : '  ') +
      label(rules).padEnd(LABEL_W - 2) +
      results.map((r) => `${r.okPages}/${r.totalPages}`.padStart(18)).join('') +
      `${ok}/${total}`.padStart(10),
  );
}
console.log('-'.repeat(LABEL_W + 18 * fixtures.length + 8));

const best = scored.reduce((a, b) => (a.ok >= b.ok ? a : b));
const winners = scored.filter((s) => s.ok === best.ok);
const current = scored.find((s) => isDefault(s.rules));
if (current === undefined) throw new Error('候选里没有包含代码当前实现的那一组');

console.log(`最优：${best.ok}/${best.total} 页，共 ${winners.length} 种组合并列`);
for (const w of winners) console.log(`  · ${label(w.rules)}`);

let failed = false;
if (current.ok !== best.ok || current.ok !== current.total) {
  const bad = current.results.find((r) => r.firstBad)?.firstBad;
  console.error(`✗ 代码里实现的那一组只对上 ${current.ok}/${current.total} 页`);
  if (bad) {
    console.error(`  第 ${bad.page + 1} 页 我们：${bad.ours.join(' | ')}`);
    console.error(`  第 ${bad.page + 1} 页 Word：${bad.theirs.join(' | ')}`);
    if (bad.deltaY !== undefined) console.error(`  文字一样，基线差 ${bad.deltaY.toFixed(3)}pt`);
  }
  failed = true;
}
// 并列就说明样本分不开这两条规则
if (winners.length > 1) {
  console.error(`✗ ${winners.length} 种组合并列最优，样本分不开 —— 页眉页脚的行数要拉开差距`);
  failed = true;
}
const worstY = Math.max(...current.results.map((r) => r.worstY));
console.log(`基线 y 的最大偏差：${worstY.toFixed(3)}pt（判据 ${TOLERANCE_PT}pt）`);
if (worstY > TOLERANCE_PT) {
  console.error(`✗ 页对上了但基线偏了 ${worstY.toFixed(3)}pt`);
  failed = true;
}

if (failed) process.exit(1);
console.log(
  [
    '',
    '✓ 页眉页脚穿刺通过：',
    '  · 页眉框顶 = w:header（到纸顶）',
    '  · 页脚量的是框底：框底 = 纸高 − w:footer',
    '  · 页边距是最小值：版心顶 = max(w:top, 页眉底)、版心底 = min(纸高 − w:bottom, 页脚顶)',
    `  三份样本 ${current.total} 页逐行文字全对（含页眉页脚与页脚里的 { PAGE }），`,
    `  基线最大偏差 ${worstY.toFixed(3)}pt。`,
  ].join('\n'),
);
