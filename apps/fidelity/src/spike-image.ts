#!/usr/bin/env node
/**
 * 图片穿刺：内嵌图在行盒里怎么摆、浮动图的八种参照框各是哪个框。
 *
 * 与 `spike:page` / `spike:header` 同一个路子 —— **不反推系数**，把整台引擎跑一遍再与真值
 * 逐行、逐图对。这里比别的穿刺多一样东西：真值里的 `images[]`（PDF 算子表里的图片落点，
 * 见 `extract-truth.ts` 的 `collectPaint`）。没有它就只能靠「图把行撑高多少」间接推，
 * 而那条路把「图摆在哪」与「行盒怎么算」两件事搅在一起。
 *
 * 三份样本：
 * - `spike-image-01`：仿宋 12pt 单倍行距，图高 4→60pt 十三档 + 22pt 字号两档 + `w:position` ±6pt 两行。
 *   回答「底边坐在基线上没有」「文字的下伸留不留」「w:position 对图片起不起作用」。
 * - `spike-image-02`：同一张图复制十三份，`positionH/V @relativeFrom` 各取一种、**偏移一律写 0**，
 *   于是量到的 x / y 就是那个框的起点本身。inside / outside 在奇偶页各放一份看镜像，
 *   `character` 排成三级阶梯看它参照的是锚点自己还是前一个字。
 * - `spike-image-03`：**开行网格**（每页 22 行 → 31.8pt）的同一条阶梯，外加倍数行距七档。
 *   01 / 02 都是关着网格量的，而中文公文一律开着网格 —— 这一份回答「含图的行吸不吸网格」，
 *   顺带把「倍数行距乘不乘在图撑起来的那一截上」照了出来（旧实现乘了，Word 没乘）。
 *
 * 这个脚本**不需要 Word**：docx 与 truth.json 都入库了，度量走随库的度量包。
 * 重新造样本才要 Windows（`pnpm truth spike-image-01 spike-image-02`）。
 *
 *   node src/spike-image.ts
 */
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDiagnosticSink, twipsToPt } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import type { DocumentLayout, ObjectRules, PageLayout } from '@uw/layout';
import { layoutDocumentWithFields, OBJECT_RULES } from '@uw/layout';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import type { TruthImage, WordTruth } from './truth-types.ts';

const APP_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
/**
 * `mode` 决定图怎么比：内嵌图比的是**它与本行基线的相对位置**，浮动图比的是纸坐标。
 * 分开不是图省事 —— 内嵌图的绝对 y 里掺着上面几十行累加出来的漂移，那是另一件事的锅。
 */
const FIXTURES = [
  { name: 'spike-image-01', mode: 'inline' as const },
  { name: 'spike-image-02', mode: 'float' as const },
  { name: 'spike-image-03', mode: 'inline' as const },
];
/** 文字基线的判据（L3） */
const TOLERANCE_PT = 0.5;
/**
 * 图片外框的判据。与文字同一条线（L4 也是 0.5pt）—— 盒高量化那条规则实现之后，
 * 44 张图的最大偏差是 0.34pt，不需要为它开口子。
 */
const IMAGE_TOLERANCE_PT = 0.5;

const CANDIDATES = {
  keepDescent: [true, false],
  raise: ['apply', 'ignore'],
  boxQuantum: ['round', 'none'],
  grid: ['apart', 'together', 'ignore'],
} as const satisfies { [K in keyof ObjectRules]: readonly ObjectRules[K][] };

interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

interface Loaded {
  name: string;
  mode: 'inline' | 'float';
  bytes: Uint8Array;
  truth: WordTruth;
}

async function load({ name, mode }: (typeof FIXTURES)[number]): Promise<Loaded> {
  const bytes = new Uint8Array(await readFile(path.join(APP_ROOT, 'fixtures', `${name}.docx`)));
  const truth = JSON.parse(
    await readFile(path.join(APP_ROOT, 'fixtures', `${name}.truth.json`), 'utf8'),
  ) as WordTruth;
  return { name, mode, bytes, truth };
}

function layoutOf(fixture: Loaded, objectRules: ObjectRules): DocumentLayout {
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
    objectRules,
  }).layout;
}

/** 一页上每一行的绝对基线（空段落在 PDF 里不落墨，真值也就没有那一行） */
function baselines(page: PageLayout): number[] {
  const out: number[] = [];
  for (const b of page.blocks) {
    if (b.kind !== 'paragraph') continue;
    for (const placed of b.lines) {
      const text = placed.line.fragments.map((f) => f.text).join('');
      if (text.trim() === '') continue;
      out.push(twipsToPt(page.geometry.content.y + placed.y + placed.line.baseline));
    }
  }
  return out.sort((a, b) => a - b);
}

/** 一页上每一张图的纸坐标外框：内嵌的从行里算，浮动的分页那一步已经算好了 */
function rects(page: PageLayout): Rect[] {
  const out: Rect[] = [];
  const g = page.geometry.content;
  for (const b of page.blocks) {
    if (b.kind !== 'paragraph') continue;
    for (const placed of b.lines) {
      for (const obj of placed.line.objects ?? []) {
        // 底边坐在基线上，raise 把它整个抬起来 —— 与渲染层同一条式子
        const bottom = g.y + placed.y + placed.line.baseline - (obj.raise ?? 0);
        out.push({
          x: twipsToPt(g.x + obj.x),
          y: twipsToPt(bottom - obj.height),
          w: twipsToPt(obj.width),
          h: twipsToPt(obj.height),
        });
      }
    }
  }
  for (const f of page.floats ?? []) {
    out.push({ x: twipsToPt(f.x), y: twipsToPt(f.y), w: twipsToPt(f.width), h: twipsToPt(f.height) });
  }
  return out.sort((a, b) => a.y - b.y || a.x - b.x);
}

const truthRects = (imgs: readonly TruthImage[] | undefined): Rect[] =>
  [...(imgs ?? [])].sort((a, b) => a.y - b.y || a.x - b.x);

interface Score {
  okLines: number;
  totalLines: number;
  okImages: number;
  totalImages: number;
  worstLine: number;
  worstImage: number;
}

/**
 * 图的可比坐标。内嵌图比的是 `x` 与**它相对本行基线的抬升**（盒底坐在基线上，
 * 抬升 = 量化 + `w:position`），浮动图比的是纸坐标 —— 内嵌图的绝对 y 里掺着上面几十行
 * 累加出来的漂移，那是 Word 内部取整的锅，不是行盒规则的锅。
 */
function comparable(rs: readonly Rect[], lines: readonly number[], mode: 'inline' | 'float'): Rect[] {
  if (mode === 'float') return [...rs];
  return rs.map((r) => {
    const bottom = r.y + r.h;
    const base = lines.reduce(
      (best, y) => (Math.abs(y - bottom) < Math.abs(best - bottom) ? y : best),
      lines[0] ?? 0,
    );
    return { ...r, y: base - bottom };
  });
}

/**
 * 打分。**行比的是逐行增量，不是累加的绝对 y** —— Word 自己的行位置带着 ±0.12pt 的抖动
 * （纯文字参照行的行距实测在 15.48–15.62pt 之间跳），几十行叠起来就能越过 L3 判据，
 * 而那是 Word 内部取整的锅，不是行盒规则的锅。行盒规则决定的恰好是**增量**：
 * 一行有多高、一张图把行撑高多少。每页第一行仍按绝对 y 比，页面几何错了照样露馅。
 */
function score(fixture: Loaded, rules: ObjectRules): Score {
  const out = layoutOf(fixture, rules);
  const s: Score = {
    okLines: 0,
    totalLines: 0,
    okImages: 0,
    totalImages: 0,
    worstLine: 0,
    worstImage: 0,
  };
  out.pages.forEach((page, i) => {
    const truth = fixture.truth.pages[i];
    if (truth === undefined) return;
    const ours = baselines(page);
    const theirs = truth.lines.map((l) => l.y);
    s.totalLines += theirs.length;
    theirs.forEach((y, k) => {
      const mine = ours[k];
      const prevMine = ours[k - 1];
      const prevTheirs = theirs[k - 1];
      if (mine === undefined) return;
      const d =
        k === 0 || prevMine === undefined || prevTheirs === undefined
          ? Math.abs(mine - y)
          : Math.abs(mine - prevMine - (y - prevTheirs));
      if (d <= TOLERANCE_PT) s.okLines += 1;
      if (d > s.worstLine) s.worstLine = d;
    });

    const theirRects = comparable(truthRects(truth.images), theirs, fixture.mode);
    const ourRects = comparable(rects(page), ours, fixture.mode);
    s.totalImages += theirRects.length;
    for (const [k, t] of theirRects.entries()) {
      const m = ourRects[k];
      if (m === undefined) continue;
      const d = Math.max(Math.abs(m.x - t.x), Math.abs(m.y - t.y));
      if (d <= IMAGE_TOLERANCE_PT) s.okImages += 1;
      if (d > s.worstImage) s.worstImage = d;
    }
  });
  return s;
}

// ── 直接测量：图底与基线差多少 ──────────────────────────────────────────────

function measureTable(fixture: Loaded): void {
  console.log(`\n① ${fixture.name} · 图底 − 同行基线（正数 = 图沉到基线以下）\n`);
  console.log('  图高pt     图底y      基线y      差pt    上伸pt（基线−图顶）');
  console.log(`  ${'-'.repeat(62)}`);
  let worst = 0;
  for (const page of fixture.truth.pages) {
    for (const im of page.images ?? []) {
      // 图所在的那一行 = 基线离图底最近的那一行（图与文字同行，差不会超过一个行高）
      const line = page.lines.reduce(
        (best, l) => (Math.abs(l.y - im.yBottom) < Math.abs(best.y - im.yBottom) ? l : best),
        page.lines[0] as { y: number },
      );
      const d = im.yBottom - line.y;
      // `w:position` 的那两行本来就该差 ±6pt，混进「最大偏差」会把结论盖掉
      if (Math.abs(d) < 3 && Math.abs(d) > Math.abs(worst)) worst = d;
      console.log(
        `  ${im.h.toFixed(1).padStart(6)} ${im.yBottom.toFixed(3).padStart(10)} ${line.y
          .toFixed(3)
          .padStart(10)} ${d.toFixed(3).padStart(8)} ${(line.y - im.y).toFixed(3).padStart(10)}`,
      );
    }
  }
  console.log(`  最大偏差 ${worst.toFixed(3)}pt（w:position 的两行除外，它们本该差 ±6pt）\n`);
}

// ── 浮动图：八种参照框各落在哪 ──────────────────────────────────────────────

function floatTable(fixture: Loaded): void {
  console.log(`\n② ${fixture.name} · 浮动图的参照框（偏移一律 0，所以量到的就是框的起点）\n`);
  const out = layoutOf(fixture, OBJECT_RULES);
  console.log('  页  图高pt      Word x/y            我们 x/y          差pt');
  console.log(`  ${'-'.repeat(64)}`);
  fixture.truth.pages.forEach((page, i) => {
    const theirs = truthRects(page.images);
    const mine = rects(out.pages[i] as PageLayout);
    theirs.forEach((im, k) => {
      const m = mine[k];
      if (m === undefined) return;
      const d = Math.max(Math.abs(m.x - im.x), Math.abs(m.y - im.y));
      console.log(
        `  ${String(i + 1).padStart(2)} ${im.h.toFixed(1).padStart(7)}  ` +
          `${im.x.toFixed(2).padStart(8)} ${im.y.toFixed(2).padStart(8)}  ` +
          `${m.x.toFixed(2).padStart(8)} ${m.y.toFixed(2).padStart(8)}  ` +
          `${d.toFixed(3).padStart(7)}${d > IMAGE_TOLERANCE_PT ? '  ✗' : ''}`,
      );
    });
  });
}

/**
 * 开网格的样本逐行看：这一行占了多高、基线离行顶多远。两个数都是**反推**的 ——
 * 真值里只有基线的绝对 y，行高只能从「下一行基线 − 本行基线」连着基线偏移一起解出来，
 * 所以这张表直接打印我们与 Word 的**相邻基线之差**，差多少一目了然。
 */
function gridTable(fixture: Loaded): void {
  console.log(`
③ ${fixture.name} · 开行网格（每页 22 行 = 31.8pt）· 相邻基线之差
`);
  const out = layoutOf(fixture, OBJECT_RULES);
  console.log('  页 行   Word 增量   我们 增量     差pt   图高pt');
  console.log(`  ${'-'.repeat(48)}`);
  fixture.truth.pages.forEach((page, i) => {
    const theirs = page.lines.map((l) => l.y);
    const ours = baselines(out.pages[i] as PageLayout);
    const imgs = truthRects(page.images);
    theirs.forEach((y, k) => {
      const mine = ours[k];
      const prevMine = ours[k - 1];
      const prevTheirs = theirs[k - 1];
      if (mine === undefined) return;
      const dTheirs = prevTheirs === undefined ? Number.NaN : y - prevTheirs;
      const dMine = prevMine === undefined ? Number.NaN : mine - prevMine;
      const d = Number.isNaN(dTheirs) ? Math.abs(mine - y) : Math.abs(dMine - dTheirs);
      const im = imgs.find((r) => Math.abs(r.y + r.h - y) < 3);
      console.log(
        `  ${String(i + 1).padStart(2)}${String(k).padStart(3)} ${(Number.isNaN(dTheirs) ? y : dTheirs)
          .toFixed(2)
          .padStart(10)} ${(Number.isNaN(dMine) ? mine : dMine).toFixed(2).padStart(10)} ${d
          .toFixed(3)
          .padStart(8)}${d > TOLERANCE_PT ? ' ✗' : '  '} ${im === undefined ? '' : im.h.toFixed(1)}`,
      );
    });
  });
}

const fixtures = await Promise.all(FIXTURES.map(load));
measureTable(fixtures[0] as Loaded);
floatTable(fixtures[1] as Loaded);
gridTable(fixtures[2] as Loaded);

const combos: ObjectRules[] = [];
for (const keepDescent of CANDIDATES.keepDescent) {
  for (const raise of CANDIDATES.raise) {
    for (const boxQuantum of CANDIDATES.boxQuantum) {
      for (const grid of CANDIDATES.grid) combos.push({ keepDescent, raise, boxQuantum, grid });
    }
  }
}
const GRID_LABEL = { apart: '两侧分算', together: '合成一个', ignore: '不吸网格' } as const;
const label = (r: ObjectRules): string =>
  `下伸 ${r.keepDescent ? '留  ' : '不留'} · w:position ${r.raise === 'apply' ? '认  ' : '不认'} · 盒高 ${
    r.boxQuantum === 'round' ? '量化到 1.5pt' : '照原样   '
  } · 网格 ${GRID_LABEL[r.grid]}`;
const isDefault = (r: ObjectRules): boolean =>
  r.keepDescent === OBJECT_RULES.keepDescent &&
  r.raise === OBJECT_RULES.raise &&
  r.boxQuantum === OBJECT_RULES.boxQuantum &&
  r.grid === OBJECT_RULES.grid;

console.log(`\n③ 行盒规则 · ${combos.length} 种组合 × ${fixtures.length} 份样本，逐行 + 逐图比对\n`);
const LABEL_W = 56;
console.log(`${'组合'.padEnd(LABEL_W)}${'行（基线）'.padStart(14)}${'图（外框）'.padStart(14)}`);
console.log('-'.repeat(LABEL_W + 28));

const scored: { rules: ObjectRules; lines: number; images: number; s: Score[] }[] = [];
for (const rules of combos) {
  const s = fixtures.map((f) => score(f, rules));
  const lines = s.reduce((n, x) => n + x.okLines, 0);
  const images = s.reduce((n, x) => n + x.okImages, 0);
  scored.push({ rules, lines, images, s });
  const totalLines = s.reduce((n, x) => n + x.totalLines, 0);
  const totalImages = s.reduce((n, x) => n + x.totalImages, 0);
  console.log(
    (isDefault(rules) ? '→ ' : '  ') +
      label(rules).padEnd(LABEL_W - 2) +
      `${lines}/${totalLines}`.padStart(14) +
      `${images}/${totalImages}`.padStart(14),
  );
}
console.log('-'.repeat(LABEL_W + 28));

const key = (x: { lines: number; images: number }): number => x.lines + x.images;
const best = scored.reduce((a, b) => (key(a) >= key(b) ? a : b));
const winners = scored.filter((x) => key(x) === key(best));
const current = scored.find((x) => isDefault(x.rules));
if (current === undefined) throw new Error('候选里没有包含代码当前实现的那一组');

console.log(`最优：${key(best)} 项对上，共 ${winners.length} 种组合并列`);
for (const w of winners) console.log(`  · ${label(w.rules)}`);

let failed = false;
const totalLines = current.s.reduce((n, x) => n + x.totalLines, 0);
const totalImages = current.s.reduce((n, x) => n + x.totalImages, 0);
if (current.lines !== totalLines || current.images !== totalImages) {
  console.error(
    `✗ 代码里实现的那一组：行 ${current.lines}/${totalLines}、图 ${current.images}/${totalImages}`,
  );
  failed = true;
}
if (winners.length > 1) {
  console.error(`✗ ${winners.length} 种组合并列最优，样本分不开这两条规则`);
  failed = true;
}
const worstLine = Math.max(...current.s.map((x) => x.worstLine));
const worstImage = Math.max(...current.s.map((x) => x.worstImage));
console.log(
  `基线最大偏差 ${worstLine.toFixed(3)}pt（判据 ${TOLERANCE_PT}）、` +
    `图外框最大偏差 ${worstImage.toFixed(3)}pt（判据 ${IMAGE_TOLERANCE_PT}）`,
);

if (failed) process.exit(1);
console.log(
  [
    '',
    '✓ 图片穿刺通过：',
    '  · 内嵌图的**盒底**坐在基线上，盒高 = 图高四舍五入到 1.5pt（图在盒里靠上放），',
    '    w:position 对它照样起作用',
    '  · 图撑的是基线以上那一截，文字自己的下伸照旧留在基线以下（行高 = 盒高 + 文字下伸）',
    '  · 浮动图：page=纸、margin/column=版心、left/topMargin 从纸边起、right/bottomMargin 从版心边起、',
    '    inside/outside 按页码奇偶镜像（**纵向镜像的是上下页边距，不是版心**）、',
    '    character=锚点前一个字的左边缘、line=行顶、paragraph=段顶',
    '  · 含图的行**参与网格吸附**（吸的是盒高 + 文字下伸），但**倍数行距不乘在图撑起来的那一截上** ——',
    '    文字侧「吸附 → 乘倍数」与对象侧「对象要的高 + 倍数多留的空白 → 吸附」各算各的，取大者',
  ].join('\n'),
);
