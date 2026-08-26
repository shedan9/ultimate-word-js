#!/usr/bin/env node
/**
 * 表格穿刺：格线的几何（占不占高、吃不吃宽）与条件格式的命中 / 层序。
 *
 * 这是**表格层的第一份真值**。在它之前，`@uw/layout` 的整个表格几何（列宽、格内边距、
 * 行高、`w:vAlign`、跨列的可用宽）与 `@uw/model` 的条件格式层序都只有规范做依据，
 * 一行都没跟 Word 比过 —— 而它们一错就是整份文档往下错位，比字体度量差半磅严重得多。
 *
 * 两份样本，各自回答一组问题：
 *
 * - `spike-table-01`：三张表、无表格样式（格式全是直接格式，于是量到的差只可能出在几何上）。
 *   六行各问一件事：默认单元格边距（108 twips 是不是真的）、`w:tcMar` 覆盖、跨列格的可用宽、
 *   `w:vAlign` 三种摆法、**6pt 粗边框吃不吃可用宽**、多段格的行高。外加表级 `w:tblCellMar`、
 *   `w:jc="center"`、`w:tblInd` 各一张表。
 * - `spike-table-02`：三张表、自定义表格样式，**每个条件设一个独一无二的字号**。
 *   于是「这一格最终几号字」= 「层序里最后一个命中它的条件是谁」，从真值的 `size` 直接读得出。
 *
 * 几何那一组按老规矩**排组合**跑（`TableRules` 的 3 × 2 = 6 种），判据是「哪一组能逐片段
 * 复现 Word」而不是残差最小。层序那一组不排组合 —— 它不是一个数而是一个全序，
 * 而**观察到的赢家只允许一个全序**（band 那一对是列压行、首末那一对是行压列、角格最后），
 * 逐格比字号本身就是判据。
 *
 * 这个脚本**不需要 Word**：docx 与 truth.json 都入库了，度量走随库的度量包。
 * 重新造样本才要 Windows（`pnpm truth spike-table-01 spike-table-02 --force`）。
 *
 *   node src/spike-table.ts
 */
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDiagnosticSink } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import type { DocumentLayout, TableRules } from '@uw/layout';
import { layoutDocumentWithFields, TABLE_RULES } from '@uw/layout';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import { piecesOf } from './flatten.ts';
import type { WordTruth } from './truth-types.ts';

const APP_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const FIXTURES = ['spike-table-01', 'spike-table-02'];

/** 文字落点的判据（L3 / L4 同一条线） */
const TOLERANCE_PT = 0.5;
/**
 * 字号的判据。真值里的 size 是从 PDF 文本矩阵反推的，实测在 ±0.05pt 抖动
 * （16pt 读成 15.96、14pt 读成 14.04），所以 0.2pt 足够分辨相邻两档整数字号。
 */
const SIZE_TOLERANCE_PT = 0.2;

const CANDIDATES = {
  gridline: ['full', 'half', 'none'],
  eatsWidth: [false, true],
} as const satisfies { [K in keyof TableRules]: readonly TableRules[K][] };

interface Loaded {
  name: string;
  bytes: Uint8Array;
  truth: WordTruth;
}

async function load(name: string): Promise<Loaded> {
  const dir = path.join(APP_ROOT, 'fixtures');
  return {
    name,
    bytes: new Uint8Array(await readFile(path.join(dir, `${name}.docx`))),
    truth: JSON.parse(await readFile(path.join(dir, `${name}.truth.json`), 'utf8')) as WordTruth,
  };
}

function layoutOf(f: Loaded, rules: TableRules): DocumentLayout {
  const sink = createDiagnosticSink();
  const loaded = loadDocument(OpcPackage.open(f.bytes), sink);
  const registry = new FontRegistry();
  for (const pack of loadBundledPacks()) registry.registerMetrics(pack);
  const measurer = createTextMeasurer(registry, {
    candidates: (family) => fontNameCandidates(loaded.fonts, family),
    diagnostics: sink,
  });
  return layoutDocumentWithFields(loaded.resolved, loaded.fields, {
    measurer,
    settings: loaded.cascade.settings,
    headerFooters: loaded.headerFooters,
    tableRules: rules,
  }).layout;
}

// ── 片段配对 ─────────────────────────────────────────────────────────────────
//
// 真值的片段与我们的片段**切法不一样**：pdf.js 按字体子集切，Word 会把一格的
// 「A1C1」拆成「A1C」+「1」，而我们按 run 切。所以不逐片段配对，而是把两边都
// 按 (y 分桶, x) 排好、把文字接起来再比 —— 与语料体检里那套分行同一个道理。
// 位置比的是**每一段连续文字的左端**，那正好是格内内容的起点，也是这份样本要量的东西。

interface Chunk {
  x: number;
  y: number;
  right: number;
  size: number;
  text: string;
}

/** 同一 y 桶、且 x 首尾相接的片段并成一段。两侧共用，切法才对得上 */
function merge(pieces: readonly Chunk[]): Chunk[] {
  const sorted = [...pieces].filter((p) => p.text.trim() !== '').sort((a, b) => a.y - b.y || a.x - b.x);
  const out: Chunk[] = [];
  for (const p of sorted) {
    const last = out[out.length - 1];
    // 同一格里被拆开的片段：基线几乎相同、x 紧挨着。y 上留 0.6pt 是取整的富余，
    // x 上留 2pt 是给字距 —— 再宽就会把相邻两格（列间至少隔着两倍单元格边距）并进来。
    if (last !== undefined && Math.abs(last.y - p.y) < 0.6 && p.x - last.right < 2) {
      last.text += p.text;
      last.right = p.right;
      continue;
    }
    out.push({ ...p });
  }
  return out;
}

function mine(doc: DocumentLayout, page: number): Chunk[] {
  const p = doc.pages[page];
  if (p === undefined) return [];
  return merge(piecesOf(p).map((q) => ({ x: q.x, y: q.y, right: q.x + q.w, size: q.size, text: q.text })));
}

function truthChunks(truth: WordTruth, page: number): Chunk[] {
  return merge(
    (truth.pages[page]?.items ?? []).map((i) => ({
      x: i.x,
      y: i.y,
      right: i.x + i.w,
      size: i.size,
      text: i.text,
    })),
  );
}

interface Score {
  name: string;
  matched: number;
  total: number;
  maxDx: number;
  maxDy: number;
  maxDsize: number;
  worst: string;
}

/**
 * y 比的是**逐段增量**而不是绝对值 —— 与 `spike-image` 同一个理由：Word 自己的行位置
 * 带着 ±0.12pt 的抖动（同一份样本里仿宋 12pt 的行距在 15.48–15.63pt 之间跳），
 * 32 段叠下来能累到 0.5pt 以上，早过了判据，而那是 Word 内部取整的锅，不是格线规则的锅。
 * 格线规则决定的恰好就是增量：跨一条格线时那一步比行距多出多少。每页第一段仍按绝对 y 比。
 *
 * x 照旧比绝对值：它不累加。
 */
function score(f: Loaded, rules: TableRules): Score {
  const doc = layoutOf(f, rules);
  let matched = 0;
  let total = 0;
  let maxDx = 0;
  let maxDy = 0;
  let maxDsize = 0;
  let worst = '';
  let worstErr = 0;

  for (let p = 0; p < f.truth.pageCount; p++) {
    const want = truthChunks(f.truth, p);
    const got = mine(doc, p);
    total += want.length;
    for (let i = 0; i < want.length; i++) {
      const w = want[i];
      if (w === undefined) continue;
      const m = got[i];
      if (m === undefined || m.text !== w.text) {
        worst = `p${p + 1} 第 ${i + 1} 段：Word「${w.text}」/ 我们「${m?.text ?? '—'}」`;
        worstErr = Number.POSITIVE_INFINITY;
        continue;
      }
      const prevW = i === 0 ? undefined : want[i - 1];
      const prevM = i === 0 ? undefined : got[i - 1];
      const dy =
        prevW === undefined || prevM === undefined
          ? Math.abs(m.y - w.y)
          : Math.abs(m.y - prevM.y - (w.y - prevW.y));
      const dx = Math.abs(m.x - w.x);
      const ds = Math.abs(m.size - w.size);
      if (dx <= TOLERANCE_PT && dy <= TOLERANCE_PT && ds <= SIZE_TOLERANCE_PT) matched++;
      else if (Math.max(dx, dy, ds) > worstErr) {
        worstErr = Math.max(dx, dy, ds);
        worst = `p${p + 1}「${w.text}」Δx=${dx.toFixed(2)} Δy=${dy.toFixed(2)} Δ字号=${ds.toFixed(2)}`;
      }
      if (dx > maxDx) maxDx = dx;
      if (dy > maxDy) maxDy = dy;
      if (ds > maxDsize) maxDsize = ds;
    }
  }
  return { name: f.name, matched, total, maxDx, maxDy, maxDsize, worst };
}

const fixtures = await Promise.all(FIXTURES.map(load));

const combos: TableRules[] = [];
for (const gridline of CANDIDATES.gridline) {
  for (const eatsWidth of CANDIDATES.eatsWidth) combos.push({ gridline, eatsWidth });
}

const label = (r: TableRules): string =>
  `格线 ${r.gridline.padEnd(4)} / 竖线${r.eatsWidth ? '吃宽' : '不吃宽'}`;
const isDefault = (r: TableRules): boolean =>
  r.gridline === TABLE_RULES.gridline && r.eatsWidth === TABLE_RULES.eatsWidth;

console.log('表格穿刺 —— 格线几何（排组合）\n');
console.log(`${''.padEnd(26)}对上 / 总数    最大 Δx    最大 Δy`);

const scored: { rules: TableRules; hit: number; total: number; s: Score[] }[] = [];
for (const rules of combos) {
  const s = fixtures.map((f) => score(f, rules));
  const hit = s.reduce((a, x) => a + x.matched, 0);
  const total = s.reduce((a, x) => a + x.total, 0);
  scored.push({ rules, hit, total, s });
  const dx = Math.max(...s.map((x) => x.maxDx));
  const dy = Math.max(...s.map((x) => x.maxDy));
  console.log(
    `${isDefault(rules) ? '→ ' : '  '}${label(rules).padEnd(24)}${`${hit} / ${total}`.padStart(11)}` +
      `${dx.toFixed(2).padStart(11)}${dy.toFixed(2).padStart(11)}`,
  );
}

const best = Math.max(...scored.map((x) => x.hit));
const winners = scored.filter((x) => x.hit === best);
console.log(`\n最优：${best} 项对上，共 ${winners.length} 种组合并列`);
for (const w of winners) console.log(`  · ${label(w.rules)}`);

const current = scored.find((x) => isDefault(x.rules));
if (current === undefined) throw new Error('实现的那一组不在候选里');

console.log('\n实现的这一组，逐份样本：');
for (const s of current.s) {
  console.log(
    `  ${s.name.padEnd(16)} ${`${s.matched} / ${s.total}`.padStart(11)}` +
      `  Δx≤${s.maxDx.toFixed(3)}  Δy≤${s.maxDy.toFixed(3)}  Δ字号≤${s.maxDsize.toFixed(3)}` +
      (s.worst === '' ? '' : `\n      最差：${s.worst}`),
  );
}

const only = winners[0];
/**
 * 对上的下限。**只许往上调**（与 `fixture.test.ts` 的 `MIN_L2_MATCH` 同一条规矩）。
 *
 * 差的那一项是 `spike-table-01` 里 `w:tcMar left="0"` 的那一格：Word 把文字放在格边
 * 右边 0.59pt 处，我们放在格边上。**它不是格线规则的锅** —— 同一列里边距 5.4pt 的
 * 那几格只差 0.32pt、边距 20pt 的差 0.24pt，三个数凑不出一条规则，没有模型就不硬凑
 * （见 `layout/src/uncalibrated.ts` 的 `TABLE_CELL_TEXT_INSET`）。
 */
const MIN_MATCH = 122;
const ok = current.hit >= MIN_MATCH && winners.length === 1 && only !== undefined && isDefault(only.rules);
console.log(
  ok
    ? `\n✓ 格线占整条线的宽（纵向）、竖线不吃可用宽是唯一满分的一组；` +
        `条件格式的命中与层序逐格对上（${current.hit} / ${current.total}，下限 ${MIN_MATCH}）`
    : '\n✗ 与真值对不上 —— 见上表',
);
if (!ok) process.exitCode = 1;
