#!/usr/bin/env node
/**
 * 表格**拆行**穿刺：一行放不下时切成两片，两片各自长什么样。
 *
 * `table-split.ts` 的文件头原来挂着四问，四条实现全是「哪种最省地方」猜的，一行都没跟
 * Word 比过。它们**不改断行**（切口只落在行间），但每一片的高度、位置、格内文字的落点
 * 全靠它们 —— 猜错一条，跨页的表格从那一页起整份文档往下错位。
 *
 * 做法与分页穿刺同一套：**把四条规则的候选排开（2⁴ = 16 组），看哪一组能逐页复现 Word**。
 * 不反推系数 —— 拆行不是一个数，是几条互相纠缠的判断（切在哪一页决定了这一片有多少地方，
 * 有多少地方又决定了切口落在第几行）。
 *
 * 样本 `spike-table-04` 七张表，每张只问一件事（详见 spec 里的 note）：
 *
 *   表甲 上下边距各 20pt，无 trHeight —— 边距在两片上怎么分、接缝那条线在哪
 *   表乙 trHeight atLeast 420pt（比版心还高）+ 表头行 —— 一片都满足不了时挪不挪走
 *   表丙 第二格 vAlign=bottom、第三格 center —— 头片认不认 vAlign
 *   表丁 trHeight atLeast 200pt + 表头行 + 2.25pt 外框 —— 富余落在哪一片、接缝取哪条边
 *   表戊 trHeight atLeast 100pt —— 表丁的第二个刻度
 *   表己 insideH 3pt —— 接缝那条线是「行自己的下边」还是「表级的下边」
 *   表庚 不拆行的 atLeast 对照 —— trHeight 量的到底是哪一段
 *
 * 判据与语料体检同构：真值的一「行」是按 y 分桶拼出来的（不认识段落也不认识表格），
 * 所以我们这边也把整页摊平再按同样的规则分桶，逐行比文字与基线 y。
 *
 * 这个脚本**不需要 Word**：docx 与 truth.json 都入库了，度量走随库的度量包。
 * 重新造样本才要 Windows（`pnpm truth spike-table-04 --force`）。
 *
 *   node src/spike-table-split.ts
 */
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDiagnosticSink } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import type { TableSplitRules } from '@uw/layout';
import { layoutDocumentWithFields, TABLE_SPLIT_RULES } from '@uw/layout';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import type { Piece } from './flatten.ts';
import { piecesOf } from './flatten.ts';
import type { WordTruth } from './truth-types.ts';

const APP_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const FIXTURE = 'spike-table-04';
/** L3 的判据 */
const TOLERANCE_PT = 0.5;
/** 与 `extract-truth.ts` 的 `groupLines()` 同一个容差（pt） */
const LINE_TOL = 1.2;

const CANDIDATES = {
  place: ['inPlace', 'nextPage'],
  margins: ['both', 'split'],
  trHeight: ['perPiece', 'wholeRow'],
  headVAlign: ['keep', 'top'],
} as const satisfies { [K in keyof TableSplitRules]: readonly TableSplitRules[K][] };

interface Line {
  y: number;
  text: string;
}

/** 按 y 分桶 → 桶内按 x 排 → 拼成一行。与 `corpus-report.ts` / `extract-truth.ts` 同构 */
function groupLines(pieces: Piece[]): Line[] {
  const sorted = [...pieces].sort((a, b) => a.y - b.y || a.x - b.x);
  const lines: Line[] = [];
  let cur: Piece[] = [];
  let curY = Number.NaN;

  const flush = (): void => {
    if (cur.length === 0) return;
    const text = [...cur]
      .sort((a, b) => a.x - b.x)
      .map((p) => p.text)
      .join('')
      .replace(/\s+/gu, '');
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

/**
 * 比的是**逐行增量**而不是绝对 y —— 与 `layout/src/image-fixture.test.ts` 同一个理由：
 * Word 自己的行距带着 ±0.12pt 的抖动（同一份文档里 15.48 / 15.60 / 15.63 都出现过），
 * 十八行叠起来能累到 0.58pt，超过 L3 的判据，而那与拆行一点关系都没有。
 * 首行仍旧比绝对 y —— 「这一片从哪儿起排」正是这份样本要量的东西。
 */
function worstStep(ours: readonly Line[], theirs: readonly Line[]): number {
  let worst = Math.abs((ours[0]?.y ?? 0) - (theirs[0]?.y ?? 0));
  for (let k = 1; k < ours.length; k++) {
    const a = (ours[k]?.y ?? 0) - (ours[k - 1]?.y ?? 0);
    const b = (theirs[k]?.y ?? 0) - (theirs[k - 1]?.y ?? 0);
    worst = Math.max(worst, Math.abs(a - b));
  }
  return worst;
}

interface Result {
  okPages: number;
  totalPages: number;
  pageCountOk: boolean;
  worstY: number;
  firstBad?: { page: number; ours: string[]; theirs: string[]; deltaY?: number };
}

const bytes = new Uint8Array(await readFile(path.join(APP_ROOT, 'fixtures', `${FIXTURE}.docx`)));
const truth = JSON.parse(
  await readFile(path.join(APP_ROOT, 'fixtures', `${FIXTURE}.truth.json`), 'utf8'),
) as WordTruth;

function run(rules: TableSplitRules): Result {
  const sink = createDiagnosticSink();
  const doc = loadDocument(OpcPackage.open(bytes), sink);
  const registry = new FontRegistry();
  for (const pack of loadBundledPacks()) registry.registerMetrics(pack);
  const measurer = createTextMeasurer(registry, {
    candidates: (family) => fontNameCandidates(doc.fonts, family),
    diagnostics: sink,
  });
  const { layout } = layoutDocumentWithFields(doc.resolved, doc.fields, {
    measurer,
    settings: doc.cascade.settings,
    headerFooters: doc.headerFooters,
    splitRules: rules,
  });

  const result: Result = {
    okPages: 0,
    totalPages: truth.pages.length,
    pageCountOk: layout.pages.length === truth.pages.length,
    worstY: 0,
  };

  layout.pages.forEach((page, i) => {
    const ours = groupLines(piecesOf(page));
    const theirs = (truth.pages[i]?.lines ?? []).map((l) => ({
      y: l.y,
      text: l.text.replace(/\s+/gu, ''),
    }));
    const sameText = ours.length === theirs.length && ours.every((l, k) => l.text === theirs[k]?.text);
    const worstY = sameText ? worstStep(ours, theirs) : 0;
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

const combos: TableSplitRules[] = [];
for (const place of CANDIDATES.place) {
  for (const margins of CANDIDATES.margins) {
    for (const trHeight of CANDIDATES.trHeight) {
      for (const headVAlign of CANDIDATES.headVAlign) {
        combos.push({ place, margins, trHeight, headVAlign });
      }
    }
  }
}

const label = (r: TableSplitRules): string =>
  `${r.place === 'inPlace' ? '就地切' : '挪走再切'} · 边距${r.margins === 'both' ? '两片各一份' : '上下分家'} · ` +
  `trHeight ${r.trHeight === 'perPiece' ? '每片一份' : '整行算'} · 头片 vAlign ${r.headVAlign}`;
const isDefault = (r: TableSplitRules): boolean =>
  (Object.keys(r) as (keyof TableSplitRules)[]).every((k) => r[k] === TABLE_SPLIT_RULES[k]);

console.log(`\n拆行规则 · ${combos.length} 种组合 × ${FIXTURE}（${truth.pages.length} 页），逐页比对\n`);
const LABEL_W = 62;
console.log(`${'组合'.padEnd(LABEL_W)}${'对上'.padStart(10)}`);
console.log('-'.repeat(LABEL_W + 10));

const scored: { rules: TableSplitRules; result: Result }[] = [];
for (const rules of combos) {
  const result = run(rules);
  scored.push({ rules, result });
  console.log(
    (isDefault(rules) ? '→ ' : '  ') +
      label(rules).padEnd(LABEL_W - 2) +
      `${result.okPages}/${result.totalPages}`.padStart(10),
  );
}
console.log('-'.repeat(LABEL_W + 10));

const best = scored.reduce((a, b) => (a.result.okPages >= b.result.okPages ? a : b));
const winners = scored.filter((s) => s.result.okPages === best.result.okPages);
const current = scored.find((s) => isDefault(s.rules));
if (current === undefined) throw new Error('候选里没有包含代码当前实现的那一组');

console.log(`最优：${best.result.okPages} / ${best.result.totalPages} 页，共 ${winners.length} 种组合并列`);
for (const w of winners) console.log(`  · ${label(w.rules)}`);

let failed = false;
if (current.result.okPages !== best.result.okPages) {
  console.error(`✗ 代码里实现的那一组只对上 ${current.result.okPages} 页，不是最优`);
  failed = true;
}
if (current.result.okPages !== current.result.totalPages) {
  const bad = current.result.firstBad;
  console.error(`✗ 代码里实现的那一组没有全对：${current.result.okPages}/${current.result.totalPages} 页`);
  if (bad) {
    console.error(`  第 ${bad.page + 1} 页 我们：${bad.ours.join(' | ')}`);
    console.error(`  第 ${bad.page + 1} 页 Word：${bad.theirs.join(' | ')}`);
    if (bad.deltaY !== undefined) console.error(`  文字一样，行距差 ${bad.deltaY.toFixed(3)}pt`);
  }
  failed = true;
}
if (winners.length > 1) {
  console.error(`✗ ${winners.length} 种组合并列最优，样本分不开这四条规则 —— 表要再加一张`);
  failed = true;
}
console.log(`首行 y 与逐行增量的最大偏差：${current.result.worstY.toFixed(3)}pt（判据 ${TOLERANCE_PT}pt）`);

if (failed) process.exit(1);
console.log(`\n✓ 「${label(current.rules)}」是唯一满分的一组（${current.result.totalPages} 页逐行全对）\n`);
