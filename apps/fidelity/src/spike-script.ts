#!/usr/bin/env node
/**
 * 脚本分桶穿刺：**纯 ASCII 的一行走哪一套行高规则**，以及**同一行里几款字体怎么合成一个行盒**。
 *
 * 与 `spike:header` / `spike:image` 同一个路子 —— 不反推系数，把整台引擎跑一遍再与真值
 * 逐页逐行对基线。理由也一样：这两问互相纠缠（合成规则错了会把「谁决定行高」的答案带偏），
 * 单独看任何一条都会被另一条污染。
 *
 * 样本 `spike-script-01`：11 页，每页四段同格式的短段连排（于是三个相邻基线差就是行高），
 * 每页换一种「`w:ascii` 槽 × `w:eastAsia` 槽」的配法，字号一律 36pt、**不开行网格**。
 * 前七页正文是**纯 ASCII**，后四页带汉字。
 *
 * 两条结论（判据见 `@uw/layout` 的 `SCRIPT_RULES` 与 `@uw/fonts` 的 `composeLineBox()`）：
 *
 * 1. **走哪一套规则看的是实际画字的那款字体**，不是这一行有没有东亚字符 ——
 *    一行「A2C6」若是等线画的，Word 照样按东亚规则算，差 30%
 * 2. **几款字体各自的行盒逐项取 max**（上取最高、下取最深），不是「取各自行高的最大值」——
 *    等线 + 宋体那一行 Word 给 50.28pt，比两款字体各自的行高都大
 *
 * 这个脚本**不需要 Word**：docx 与 truth.json 都入库了，度量走随库的度量包。
 * 重新造样本才要 Windows（`pnpm truth spike-script-01`）。
 *
 *   node src/spike-script.ts
 */
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDiagnosticSink, twipsToPt } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import type { DocumentLayout, ScriptRules } from '@uw/layout';
import { layoutDocumentWithFields, SCRIPT_RULES } from '@uw/layout';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import type { WordTruth } from './truth-types.ts';

const APP_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const FIXTURES = ['spike-script-01'];
/** L3 的判据。四种说法在这份样本上相差 3.5–14pt，0.5pt 分得干干净净 */
const TOLERANCE_PT = 0.5;

const CANDIDATES = {
  eastAsianBy: ['font', 'line'],
  compose: ['maxSides', 'maxHeight'],
  box: ['all', 'eastAsiaBucket'],
} as const satisfies { [K in keyof ScriptRules]: readonly ScriptRules[K][] };

interface Line {
  y: number;
  text: string;
}

interface Result {
  okPages: number;
  totalPages: number;
  worstY: number;
  firstBad?: { page: number; ours: Line[]; theirs: Line[]; deltaY?: number };
}

/**
 * 比文字时把空白全去掉。真值里的「P8 汉 TimesAsciiDengEa」多出来的两个空格是
 * **中西文自动间距**（1/8 em）在 PDF 里张开的缝 —— 抽真值那一步按片段间距补空格，
 * 它是宽度那一维的事。这份样本量的是基线 y，留着它只会让四页永远判「文字不一样」。
 */
const squash = (s: string): string => s.replace(/\s+/gu, '');

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

function layoutOf(fixture: Loaded, scriptRules: ScriptRules): DocumentLayout {
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
    scriptRules,
  }).layout;
}

function ourLines(layout: DocumentLayout, pageIndex: number): Line[] {
  const page = layout.pages[pageIndex];
  if (page === undefined) return [];
  const out: Line[] = [];
  for (const b of page.blocks) {
    if (b.kind !== 'paragraph') continue;
    for (const placed of b.lines) {
      const text = placed.line.fragments
        .map((f) => f.text)
        .join('')
        .replace(/\s+$/u, '');
      if (text !== '')
        out.push({ y: twipsToPt(page.geometry.content.y + placed.y + placed.line.baseline), text });
    }
  }
  return out;
}

function run(fixture: Loaded, rules: ScriptRules): Result {
  const layout = layoutOf(fixture, rules);
  const result: Result = { okPages: 0, totalPages: fixture.truth.pages.length, worstY: 0 };
  fixture.truth.pages.forEach((truthPage, i) => {
    const ours = ourLines(layout, i);
    const theirs = truthPage.lines.map((l) => ({ y: l.y, text: l.text }));
    const sameText =
      ours.length === theirs.length && ours.every((l, k) => squash(l.text) === squash(theirs[k]?.text ?? ''));
    const worstY = sameText ? Math.max(0, ...ours.map((l, k) => Math.abs(l.y - (theirs[k]?.y ?? l.y)))) : 0;
    if (!sameText || worstY > TOLERANCE_PT) {
      result.firstBad ??= { page: i, ours, theirs, ...(sameText ? { deltaY: worstY } : {}) };
      return;
    }
    result.okPages += 1;
    if (worstY > result.worstY) result.worstY = worstY;
  });
  return result;
}

const fixtures = await Promise.all(FIXTURES.map(load));

const combos: ScriptRules[] = [];
for (const eastAsianBy of CANDIDATES.eastAsianBy) {
  for (const compose of CANDIDATES.compose) {
    for (const box of CANDIDATES.box) combos.push({ eastAsianBy, compose, box });
  }
}

const label = (r: ScriptRules): string =>
  `按 ${r.eastAsianBy.padEnd(4)} 判 · 合成 ${r.compose.padEnd(9)} · 行盒 ${r.box}`;
const isDefault = (r: ScriptRules): boolean =>
  r.eastAsianBy === SCRIPT_RULES.eastAsianBy &&
  r.compose === SCRIPT_RULES.compose &&
  r.box === SCRIPT_RULES.box;

console.log(`\n脚本分桶与行盒合成 · ${combos.length} 种组合 × ${fixtures.length} 份样本，逐页比对\n`);
const LABEL_W = 48;
console.log(`${'组合'.padEnd(LABEL_W)}${'对上的页'.padStart(12)}`);
console.log('-'.repeat(LABEL_W + 12));

const scored: { rules: ScriptRules; ok: number; total: number; results: Result[] }[] = [];
for (const rules of combos) {
  const results = fixtures.map((f) => run(f, rules));
  const ok = results.reduce((n, r) => n + r.okPages, 0);
  const total = results.reduce((n, r) => n + r.totalPages, 0);
  scored.push({ rules, ok, total, results });
  console.log(
    (isDefault(rules) ? '→ ' : '  ') + label(rules).padEnd(LABEL_W - 2) + `${ok}/${total}`.padStart(12),
  );
}
console.log('-'.repeat(LABEL_W + 12));

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
    console.error(
      `  第 ${bad.page + 1} 页 我们：${bad.ours.map((l) => `${l.text}@${l.y.toFixed(2)}`).join(' | ')}`,
    );
    console.error(
      `  第 ${bad.page + 1} 页 Word：${bad.theirs.map((l) => `${l.text}@${l.y.toFixed(2)}`).join(' | ')}`,
    );
    if (bad.deltaY !== undefined) console.error(`  文字一样，基线差 ${bad.deltaY.toFixed(3)}pt`);
  }
  failed = true;
}
if (winners.length > 1) {
  console.error(`✗ ${winners.length} 种组合并列最优，样本分不开 —— 要再加互相分岔的配法`);
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
    '✓ 脚本分桶穿刺通过：',
    '  · 行高走东亚规则还是拉丁规则，看的是**实际画字的那款字体**，逐段判',
    '    —— 不是行里有没有东亚字符，不是 w:eastAsia 槽，也不是 w:hint',
    '  · 同一行里几款字体，各自的行盒**逐项取 max**（上取最高、下取最深）',
    '    —— 不是「取各自行高的最大值」，那个数比 Word 小 1.5pt',
    `  一份样本 ${current.total} 页逐行全对，基线最大偏差 ${worstY.toFixed(3)}pt。`,
  ].join('\n'),
);
