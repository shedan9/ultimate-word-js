#!/usr/bin/env node
/**
 * 语料体检：拿真实文档走完整条链（解包 → 级联 → 度量 → 断行 → 分页 → 域求值），
 * 与它的 `*.truth.json` 比 L0（页数）与 L2（每行的文字），并把诊断按 code 汇总。
 *
 * **它不是测试**。新语料进来时页数对不上、诊断一大把是常态，这个脚本的用处是
 * 「差在哪、差多少」一眼看清，再决定哪几条值得写成断言。真正的闸门在
 * `layout/src/fixture.test.ts`（那里的 `MIN_L2_MATCH` 只许往上调）。
 *
 * ## 为什么要自己重做一遍「分行」
 *
 * 真值里的一「行」是 `extract-truth.ts` 的 `groupLines()` 按 **y 分桶**分出来的：
 * 同一个 y（容差内）的所有文字片段按 x 排好、拼成一串。它**不认识段落，也不认识表格** ——
 * 一行两个单元格的文字会被拼成同一行。所以我们这边要比得上，就必须走同一条路：
 * 把整页所有文字（正文段落 + **表格单元格** + **页眉页脚**）摊平成「绝对坐标 + 文字」
 * 的片段，再按同样的规则分桶。
 *
 * 只收正文段落的话，表格类文档会得出「我们 234 行 / Word 505 行」这种毫无意义的数字 ——
 * 这不是保真度差，是根本没在比同一样东西。第一版就是这么写的，记在这儿免得再写一次。
 *
 * 用法：`node src/corpus-report.ts [name ...]`（不给名字就跑 fixtures 里所有非 spike 的）
 */
import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDiagnosticSink } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import { layoutDocumentWithFields } from '@uw/layout';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import type { Piece } from './flatten.ts';
import { piecesOf } from './flatten.ts';

const FIXTURES = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'fixtures');

/** 与 `extract-truth.ts` 的 `groupLines()` 同一个容差（pt） */
const LINE_TOL = 1.2;

interface TruthFile {
  pageCount: number;
  pages: { lines: { text: string; y: number; x: number }[] }[];
}

const trim = (s: string): string => s.replace(/\s+$/u, '');

/** 按 y 分桶 → 桶内按 x 排 → 拼成一行。与 `extract-truth.ts` 的 `groupLines()` 同构 */
function groupLines(pieces: Piece[]): { y: number; x: number; text: string }[] {
  const sorted = [...pieces].sort((a, b) => a.y - b.y || a.x - b.x);
  const lines: { y: number; x: number; text: string }[] = [];
  let cur: Piece[] = [];
  let curY = Number.NaN;

  const flush = (): void => {
    if (cur.length === 0) return;
    const byX = [...cur].sort((a, b) => a.x - b.x);
    const text = trim(byX.map((p) => p.text).join(''));
    if (text !== '') lines.push({ y: curY, x: byX[0]?.x ?? 0, text });
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

// ── 报告 ──────────────────────────────────────────────────────────────────────

function names(): string[] {
  const args = process.argv.slice(2).filter((a) => !a.startsWith('--'));
  if (args.length > 0) return args;
  return readdirSync(FIXTURES)
    .filter((f) => f.endsWith('.docx') && !f.startsWith('spike-') && !f.startsWith('~$'))
    .map((f) => path.basename(f, '.docx'))
    .sort();
}

const verbose = process.argv.includes('--diff');

for (const name of names()) {
  const sink = createDiagnosticSink();
  console.log(`\n${'═'.repeat(72)}\n${name}\n${'═'.repeat(72)}`);

  const bytes = new Uint8Array(readFileSync(path.join(FIXTURES, `${name}.docx`)));
  const doc = loadDocument(OpcPackage.open(bytes), sink);
  const registry = new FontRegistry();
  for (const pack of loadBundledPacks()) registry.registerMetrics(pack);
  const measurer = createTextMeasurer(registry, {
    candidates: (family) => fontNameCandidates(doc.fonts, family),
    diagnostics: sink,
  });

  const { layout, passes, converged } = layoutDocumentWithFields(doc.resolved, doc.fields, {
    measurer,
    settings: doc.cascade.settings,
    diagnostics: sink,
    headerFooters: doc.headerFooters,
  });

  const truth = JSON.parse(readFileSync(path.join(FIXTURES, `${name}.truth.json`), 'utf8')) as TruthFile;
  const ourPages = layout.pages.map((p) => groupLines(piecesOf(p)));
  const theirPages = truth.pages.map((p) =>
    p.lines.map((l) => ({ ...l, text: trim(l.text) })).filter((l) => l.text !== ''),
  );

  const same = layout.pages.length === truth.pageCount;
  console.log(`L0 页数：我们 ${layout.pages.length} / Word ${truth.pageCount}  ${same ? '✅' : '❌'}`);
  console.log(`   每页行数 我们：${ourPages.map((p) => p.length).join(' ')}`);
  console.log(`            Word：${theirPages.map((p) => p.length).join(' ')}`);
  if (passes > 1 || !converged) console.log(`   域求值 ${passes} 趟${converged ? '收敛' : '**没收敛**'}`);

  const ours = ourPages.flat();
  const theirs = theirPages.flat();
  const n = Math.min(ours.length, theirs.length);
  let matched = 0;
  const diffs: number[] = [];
  for (let i = 0; i < n; i++) {
    if (ours[i]?.text === theirs[i]?.text) matched++;
    else diffs.push(i);
  }
  const pct = theirs.length === 0 ? 0 : (matched / theirs.length) * 100;
  console.log(
    `L2 断行：${matched} / ${theirs.length} 行逐字一致（${pct.toFixed(1)}%，我们排出 ${ours.length} 行）`,
  );
  for (const i of diffs.slice(0, verbose ? 40 : 3)) {
    console.log(`   第 ${i} 行  我们「${ours[i]?.text}」`);
    console.log(`   ${' '.repeat(String(i).length)}      Word「${theirs[i]?.text}」`);
  }
  if (!verbose && diffs.length > 3) console.log(`   …… 还有 ${diffs.length - 3} 处（加 --diff 全看）`);

  const byCode = new Map<string, number>();
  for (const d of sink.list()) byCode.set(d.code, (byCode.get(d.code) ?? 0) + 1);
  if (byCode.size === 0) console.log('诊断：一条都没有 ✅');
  else {
    console.log(`诊断：${sink.list().length} 条`);
    for (const [code, count] of [...byCode].sort((a, b) => b[1] - a[1])) {
      console.log(`   ${code} × ${count}  例：${sink.list().find((d) => d.code === code)?.message ?? ''}`);
    }
  }
}
