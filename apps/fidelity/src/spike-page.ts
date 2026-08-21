#!/usr/bin/env node
/**
 * 分页穿刺：孤行寡行的下限、keepNext 的接缝、页首的段前间距。
 *
 * 与其他几个穿刺不同，这一个**不反推系数**，而是把整台引擎跑一遍再与真值逐页对 ——
 * 分页规则不是一个数，是一组互相纠缠的判断（孤行寡行会先一步把段落推走，
 * keepNext 的接缝又得看下一段肯不肯拆），单独反推任何一条都会被另一条污染。
 * 所以做法是：**把三条规则各自的候选组合排开，看哪一组能逐页复现 Word**。
 *
 * 样本（`spike-page-01/02`）的版心刻意做成「一页恰好 11 行、一行 18 个汉字、固定行距 20pt」，
 * 于是行高与字宽都不依赖任何待标定的度量，阶梯靠垫行的条数移动断页点。
 *
 * 三条结论（判据见 `@uw/layout` 的 `PAGINATION_RULES`）：
 *
 * 1. 孤行寡行保底 **2 行**：垫 7 行时自然断点是 4|1，Word 给的是 3|2
 * 2. 段前间距落在页首 **不算**：24pt 段前的段落被顶到页首时，首行基线与其余每一页一样
 * 3. keepNext 的接缝要留出下一块 **最少能放多少**，而不是它的第一行：下一段只有 2 行时
 *    孤行寡行不许它拆，于是它整块都得跟着走
 *
 * 这个脚本**不需要 Word**：docx 与 truth.json 都入库了，度量走随库的度量包。
 * 重新造样本才要 Windows（`pnpm truth spike-page-01 spike-page-02`）。
 *
 *   node src/spike-page.ts
 */
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDiagnosticSink, twipsToPt } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import type { PaginationRules } from '@uw/layout';
import { layoutDocument, PAGINATION_RULES } from '@uw/layout';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import type { WordTruth } from './truth-types.ts';

const APP_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const FIXTURES = ['spike-page-01', 'spike-page-02'];
/** L3 的判据。这里顺带把基线也对一遍 —— 分页错了页会对不上，基线错了 y 会对不上 */
const TOLERANCE_PT = 0.5;

/** 候选：每条规则的可能取值。笛卡尔积全跑一遍 */
const CANDIDATES = {
  widowMinLines: [1, 2, 3],
  spaceBeforeAtPageTop: [false, true],
  keepNextJoin: ['min-chunk', 'first-line', 'whole-block'],
} as const satisfies { [K in keyof PaginationRules]: readonly PaginationRules[K][] };

interface Result {
  /** 与真值逐行文字一致的页数 */
  okPages: number;
  totalPages: number;
  /** 我们排出来的页数与 Word 的页数是否一致 */
  pageCountOk: boolean;
  /** 基线 y 的最大偏差（pt），只统计对上的那些页 */
  worstY: number;
  /** 第一处对不上的页，报错时直接能看 */
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

function run(fixture: Loaded, rules: PaginationRules): Result {
  const sink = createDiagnosticSink();
  const doc = loadDocument(OpcPackage.open(fixture.bytes), sink);
  const registry = new FontRegistry();
  for (const pack of loadBundledPacks()) registry.registerMetrics(pack);
  const measurer = createTextMeasurer(registry, {
    candidates: (family) => fontNameCandidates(doc.fonts, family),
    diagnostics: sink,
  });
  const out = layoutDocument(doc.resolved, { measurer, settings: doc.cascade.settings, rules });

  const result: Result = {
    okPages: 0,
    totalPages: fixture.truth.pages.length,
    pageCountOk: out.pages.length === fixture.truth.pages.length,
    worstY: 0,
  };

  out.pages.forEach((page, i) => {
    // 空段落在 PDF 里不落墨，真值也就没有那一行 —— 比对前要先滤掉，否则整页错位
    const ours = page.blocks
      .filter((b) => b.kind === 'paragraph')
      .flatMap((b) => b.lines)
      .map((placed) => ({
        y: twipsToPt(page.geometry.content.y + placed.y + placed.line.baseline),
        text: placed.line.fragments
          .map((f) => f.text)
          .join('')
          .replace(/\s+$/u, ''),
      }))
      .filter((l) => l.text !== '');
    const theirs = fixture.truth.pages[i]?.lines ?? [];

    // 「这一页对上了」= 行文字一致**且**每一行的基线都在判据内。
    // 光比文字分不开「页首的段前间距算不算」：那 24pt 不改任何一行落在哪一页，
    // 只把整页文字往下推 24pt —— 不比 y 的话它就成了一个测不出来的自由参数
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

/** 笛卡尔积。规则只有三条，写死比通用组合器好读 */
const combos: PaginationRules[] = [];
for (const widowMinLines of CANDIDATES.widowMinLines) {
  for (const spaceBeforeAtPageTop of CANDIDATES.spaceBeforeAtPageTop) {
    for (const keepNextJoin of CANDIDATES.keepNextJoin) {
      combos.push({ widowMinLines, spaceBeforeAtPageTop, keepNextJoin });
    }
  }
}

const label = (r: PaginationRules): string =>
  `寡行下限 ${r.widowMinLines} · 页首段前 ${r.spaceBeforeAtPageTop ? '加' : '不加'} · 接缝 ${r.keepNextJoin}`;
const isDefault = (r: PaginationRules): boolean =>
  r.widowMinLines === PAGINATION_RULES.widowMinLines &&
  r.spaceBeforeAtPageTop === PAGINATION_RULES.spaceBeforeAtPageTop &&
  r.keepNextJoin === PAGINATION_RULES.keepNextJoin;

console.log(`\n分页规则 · ${combos.length} 种组合 × ${fixtures.length} 份样本，逐页比对\n`);
const LABEL_W = 46;
const header = fixtures.map((f) => f.name.padStart(18)).join('');
console.log(`${'组合'.padEnd(LABEL_W)}${header}   总对上`);
console.log('-'.repeat(LABEL_W + 18 * fixtures.length + 8));

const scored: { rules: PaginationRules; ok: number; total: number; results: Result[] }[] = [];
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
if (current.ok !== best.ok) {
  console.error(`✗ 代码里实现的那一组只对上 ${current.ok}/${current.total} 页，不是最优`);
  failed = true;
}
if (current.ok !== current.total) {
  const bad = current.results.find((r) => r.firstBad)?.firstBad;
  console.error(`✗ 代码里实现的那一组没有全对：${current.ok}/${current.total} 页`);
  if (bad) {
    console.error(`  第 ${bad.page + 1} 页 我们：${bad.ours.join(' | ')}`);
    console.error(`  第 ${bad.page + 1} 页 Word：${bad.theirs.join(' | ')}`);
    if (bad.deltaY !== undefined) console.error(`  文字一样，基线差 ${bad.deltaY.toFixed(3)}pt`);
  }
  failed = true;
}
// 并列就说明样本分不开 —— 与基线穿刺的 MIN_MARGIN 是同一个判据，只是这里没有连续量可比
if (winners.length > 1) {
  console.error(`✗ ${winners.length} 种组合并列最优，样本分不开这三条规则 —— 阶梯要加密`);
  failed = true;
}
const worstY = Math.max(...current.results.map((r) => r.worstY));
console.log(`基线 y 的最大偏差：${worstY.toFixed(3)}pt（判据 ${TOLERANCE_PT}pt）`);
if (worstY > TOLERANCE_PT) {
  console.error(`✗ 页对上了但基线偏了 ${worstY.toFixed(3)}pt —— 那是行盒的问题，不是分页的`);
  failed = true;
}

if (failed) process.exit(1);
console.log(
  [
    '',
    '✓ 分页穿刺通过：',
    '  · 孤行寡行保底 2 行（页底至少留 2 行，下一页至少接 2 行）',
    '  · 段前间距落在页首不算',
    '  · keepNext 的接缝要留出下一块「最少能放多少」，不是它的第一行',
    `  两份样本 ${current.total} 页逐行文字全对，基线最大偏差 ${worstY.toFixed(3)}pt。`,
  ].join('\n'),
);
