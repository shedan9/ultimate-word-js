#!/usr/bin/env node
/**
 * 分桶穿刺的**宽度那一半**：歧义字符与中性字符进哪个桶、中西文自动间距有多宽。
 * （行高那一半是 `spike:script`，两者共用 `spike-script-01` / `spike-width-01` 这一对样本。）
 *
 * 与别的 spike 同一个路子 —— 不反推系数，把整台引擎跑一遍再与真值逐行对。
 * 只是这一份对的是**横向**：每一行比两样东西，
 *
 * 1. **片段的字体序列**（相邻重复的合并掉）。真值的 `TruthItem.font` 直接说出 Word
 *    用哪款字体画了这个字 —— 字体一换，PDF 里就换一次 `Tf`、起一个新片段。
 *    于是「§ 进了哪个桶」是**读**出来的而不是从宽度反推的
 * 2. **行末 x**。分桶错了宽度就错（宋体的 `§` 是 1 em、Times 的只有 0.5 em），
 *    自动间距错了每个中西文边界都错 —— 两种错都落在这一个数上
 *
 * 唯一读不得字体名的是**空格**：Word 画它时不换 `Tf`，PDF 里它跟着前一个字走，
 * 而推进宽度才是另一款字体的（见 `@uw/fonts` 的 `neutralTakesEastAsia`）。
 * 空格那几段靠第 2 项分辨 —— 0.5 em 与 0.25 em 差 9pt / 36pt 字，远大于判据。
 *
 * 这个脚本**不需要 Word**：docx 与 truth.json 都入库了，度量走随库的度量包。
 * 重新造样本才要 Windows（`pnpm truth spike-width-01`）。
 *
 *   node src/spike-width.ts
 */
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDiagnosticSink, twipsToPt } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import type { DocumentLayout, WidthRules } from '@uw/layout';
import { layoutDocumentWithFields, WIDTH_RULES } from '@uw/layout';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import type { WordTruth } from './truth-types.ts';

const APP_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const FIXTURES = ['spike-width-01'];
/**
 * L4 的判据。这份样本里各条假设互相差 9–18pt（半个字到一整个字），0.5pt 分得干干净净；
 * 真值本身的坐标噪声在 0.15pt 量级（Times 的片段宽度是逐字形加出来再取整的）。
 */
const TOLERANCE_PT = 0.5;

/**
 * 文档里的字体名 → 真值里的 PostScript 名。
 *
 * 写死在这里而不是从 `fontTable.xml` 推：这一份样本只用两款字体，而「本地化名 →
 * 磁盘字体 → PostScript 名」那条链本身是另一件要标定的事（`fontNameCandidates`），
 * 混进来会让读数依赖一个未标定的东西。
 */
const PS_NAME: Record<string, string> = {
  宋体: 'SimSun',
  'Times New Roman': 'TimesNewRomanPSMT',
};

const CANDIDATES = {
  ambiguous: ['hint', 'eastAsia', 'latin'],
  neutral: ['either', 'eitherHinted', 'prev', 'none'],
  autoSpaceEm: [0.25, 0.125, 0],
  autoSpaceSize: ['prev', 'eastAsia'],
  autoSpaceScope: ['eastAsianCp', 'bucket'],
} as const satisfies { [K in keyof WidthRules]: readonly WidthRules[K][] };

interface Line {
  /** 行末 x，pt，纸左边缘起 */
  xEnd: number;
  text: string;
  /** 片段字体序列，相邻重复已合并 */
  fonts: string[];
}

interface Result {
  okLines: number;
  totalLines: number;
  worstX: number;
  firstBad?: { page: number; row: number; ours?: Line; theirs: Line };
}

/** 比文字时把空白全去掉：真值里的空格既有真的、也有 pdf.js 为缝隙补出来的 */
const squash = (s: string): string => s.replace(/\s+/gu, '');

/** 相邻重复的字体合并掉 —— 我们按 run 切片段，Word 按 `Tf` 切，粒度本来就不同 */
function dedupe(fonts: readonly string[]): string[] {
  const out: string[] = [];
  for (const f of fonts) {
    if (out[out.length - 1] !== f) out.push(f);
  }
  return out;
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

function layoutOf(fixture: Loaded, widthRules: WidthRules): DocumentLayout {
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
    widthRules,
  }).layout;
}

function ourLines(layout: DocumentLayout, pageIndex: number): Line[] {
  const page = layout.pages[pageIndex];
  if (page === undefined) return [];
  const left = page.geometry.content.x;
  const out: Line[] = [];
  for (const b of page.blocks) {
    if (b.kind !== 'paragraph') continue;
    for (const placed of b.lines) {
      const frags = placed.line.fragments.filter((f) => f.text !== '');
      if (frags.length === 0) continue;
      const text = frags
        .map((f) => f.text)
        .join('')
        .replace(/\s+$/u, '');
      if (text === '') continue;
      const xEnd = Math.max(...frags.map((f) => f.x + f.width));
      out.push({
        xEnd: twipsToPt(left + xEnd),
        text,
        // 行尾空格不参与：Word 画不画它都不影响行宽，而我们的片段里它还在
        fonts: dedupe(frags.filter((f) => f.text.trim() !== '').map((f) => PS_NAME[f.font] ?? f.font)),
      });
    }
  }
  return out;
}

function truthLines(fixture: Loaded, pageIndex: number): Line[] {
  const page = fixture.truth.pages[pageIndex];
  if (page === undefined) return [];
  return page.lines.map((l) => ({
    xEnd: l.xEnd,
    text: l.text,
    fonts: dedupe(
      l.items
        .map((i) => page.items[i])
        .filter((it) => it !== undefined && it.text.trim() !== '')
        .map((it) => (it as { font: string }).font),
    ),
  }));
}

function run(fixture: Loaded, rules: WidthRules): Result {
  const layout = layoutOf(fixture, rules);
  const result: Result = { okLines: 0, totalLines: 0, worstX: 0 };
  fixture.truth.pages.forEach((_, i) => {
    const ours = ourLines(layout, i);
    const theirs = truthLines(fixture, i);
    result.totalLines += theirs.length;
    theirs.forEach((t, k) => {
      const o = ours[k];
      const same =
        o !== undefined &&
        squash(o.text) === squash(t.text) &&
        o.fonts.join('>') === t.fonts.join('>') &&
        Math.abs(o.xEnd - t.xEnd) <= TOLERANCE_PT;
      if (!same) {
        result.firstBad ??= { page: i, row: k, ...(o === undefined ? {} : { ours: o }), theirs: t };
        return;
      }
      result.okLines += 1;
      result.worstX = Math.max(result.worstX, Math.abs(o.xEnd - t.xEnd));
    });
  });
  return result;
}

const fixtures = await Promise.all(FIXTURES.map(load));

const combos: WidthRules[] = [];
for (const ambiguous of CANDIDATES.ambiguous) {
  for (const neutral of CANDIDATES.neutral) {
    for (const autoSpaceEm of CANDIDATES.autoSpaceEm) {
      for (const autoSpaceSize of CANDIDATES.autoSpaceSize) {
        for (const autoSpaceScope of CANDIDATES.autoSpaceScope) {
          combos.push({ ambiguous, neutral, autoSpaceEm, autoSpaceSize, autoSpaceScope });
        }
      }
    }
  }
}

const label = (r: WidthRules): string =>
  `歧义 ${r.ambiguous.padEnd(8)} · 中性 ${r.neutral.padEnd(12)} · 间距 ${String(r.autoSpaceEm).padEnd(5)}` +
  ` · 按 ${r.autoSpaceSize.padEnd(8)} · 范围 ${r.autoSpaceScope}`;
const isDefault = (r: WidthRules): boolean =>
  (Object.keys(WIDTH_RULES) as (keyof WidthRules)[]).every((k) => r[k] === WIDTH_RULES[k]);

console.log(`\n宽度分桶与中西文间距 · ${combos.length} 种组合 × ${fixtures.length} 份样本，逐行比对\n`);
const LABEL_W = 74;
console.log(`${'组合'.padEnd(LABEL_W)}${'对上的行'.padStart(12)}`);
console.log('-'.repeat(LABEL_W + 12));

const scored: { rules: WidthRules; ok: number; total: number; results: Result[] }[] = [];
for (const rules of combos) {
  const results = fixtures.map((f) => run(f, rules));
  const ok = results.reduce((n, r) => n + r.okLines, 0);
  const total = results.reduce((n, r) => n + r.totalLines, 0);
  scored.push({ rules, ok, total, results });
}
// 只打印前十名与代码当前那一组 —— 144 行看不出名堂
const ranked = [...scored].sort((a, b) => b.ok - a.ok);
for (const s of ranked.slice(0, 10)) {
  console.log(
    (isDefault(s.rules) ? '→ ' : '  ') +
      label(s.rules).padEnd(LABEL_W - 2) +
      `${s.ok}/${s.total}`.padStart(12),
  );
}
const currentRank = ranked.findIndex((s) => isDefault(s.rules));
if (currentRank >= 10) {
  const s = ranked[currentRank] as (typeof ranked)[number];
  console.log(`  …\n→ ` + label(s.rules).padEnd(LABEL_W - 2) + `${s.ok}/${s.total}`.padStart(12));
}
console.log('-'.repeat(LABEL_W + 12));

const best = ranked[0] as (typeof ranked)[number];
const winners = scored.filter((s) => s.ok === best.ok);
const current = scored.find((s) => isDefault(s.rules));
if (current === undefined) throw new Error('候选里没有包含代码当前实现的那一组');

console.log(`最优：${best.ok}/${best.total} 行，共 ${winners.length} 种组合并列`);
for (const w of winners) console.log(`  · ${label(w.rules)}`);

let failed = false;
if (current.ok !== best.ok || current.ok !== current.total) {
  const bad = current.results.find((r) => r.firstBad)?.firstBad;
  console.error(`✗ 代码里实现的那一组只对上 ${current.ok}/${current.total} 行`);
  if (bad) {
    const fmt = (l: Line | undefined): string =>
      l === undefined
        ? '（这一行我们根本没排出来）'
        : `${JSON.stringify(l.text)} 末端 ${l.xEnd.toFixed(2)} 字体 ${l.fonts.join('>')}`;
    console.error(`  第 ${bad.page + 1} 页第 ${bad.row + 1} 行`);
    console.error(`    我们：${fmt(bad.ours)}`);
    console.error(`    Word：${fmt(bad.theirs)}`);
  }
  failed = true;
}
if (winners.length > 1) {
  console.error(`✗ ${winners.length} 种组合并列最优，样本分不开 —— 要再加互相分岔的段落`);
  failed = true;
}
const worstX = Math.max(...current.results.map((r) => r.worstX));
console.log(`行末 x 的最大偏差：${worstX.toFixed(3)}pt（判据 ${TOLERANCE_PT}pt）`);
if (worstX > TOLERANCE_PT) {
  console.error(`✗ 行对上了但末端偏了 ${worstX.toFixed(3)}pt`);
  failed = true;
}

if (failed) process.exit(1);
console.log(
  [
    '',
    '✓ 宽度分桶穿刺通过：',
    '  · 歧义字符（EastAsianWidth = A）跟着 w:hint 走，与邻居无关',
    '  · 空格随东亚邻居（任一侧），**与 hint 无关**；/ 与 - 这类中性字符不随',
    '  · 中西文自动间距是 1/4 em（不是 1/8），按**接缝前面**那个字符的字号算',
    '  · 靠 hint 才进东亚桶的歧义字符**不加**自动间距（与全角标点同理）',
    `  一份样本 ${current.total} 行逐行全对，行末 x 最大偏差 ${worstX.toFixed(3)}pt。`,
  ].join('\n'),
);
