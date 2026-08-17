#!/usr/bin/env node
/**
 * 基线穿刺：基线在行高里的位置。
 *
 * Phase 0 只钉死了行高的**总量**（含东亚文字的行 = win 跨度 × 1.3）。总量对了不等于字画对了 ——
 * 基线偏 3pt，整页文字会整体上移，而每行的间距看起来完全正常。行盒装配、分页、DOM 渲染
 * 全都等这一问，所以它是 DEVELOPMENT-PLAN §13 的第一条。
 *
 * 做法：fixture 里每段都用 `pageBreakBefore` 顶到自己那页的最上面，于是
 * 「首行基线 − 版心顶」就是基线在行盒里的位置，与前面排了什么无关。拿它和四个候选假设比 ——
 *
 *   below  额外行距全在基线以下（= 基线就落在 winAscent 处）
 *   half   上下均分
 *   prop   按 ascent / descent 的比例分（等价于 ascent 也乘 1.3）
 *   above  全在基线以上
 *
 * 东亚行与拉丁行**分开评**，因为它们的答案不一样（这本身就是结论之一）：
 * 东亚那 30% 上下均分，拉丁的 GDI 外部行距整块在基线以上。混在一起评会得到一个
 * 两边都不满意的中间值，还会让「拉丁那条规则」这个发现整个消失。
 *
 * 判据不是「哪个看起来对」，而是**每个假设各自的最大误差**：赢的那个必须比次优好几倍，
 * 否则说明样本分辨力不够 —— 这正是 spike-baseline-02 把字号放大到 72pt 的原因，
 * 12pt 下拉丁一侧 half 与 above 只差 0.26pt，与坐标噪声同量级。
 *
 *   node src/spike-baseline.ts            # 跑三份 fixture
 *   node src/spike-baseline.ts <name>...  # 指定 fixture
 */
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { ptToTwips, twipsToPt } from '@uw/core';
import { lineMetrics } from '@uw/fonts';
import { assertWindows } from './platform.ts';
import { hasEastAsianText, loadSpikeFont } from './spike-fonts.ts';
import type { TruthPage, WordTruth } from './truth-types.ts';

assertWindows({ tool: '基线穿刺', needs: '要读 C:/Windows/Fonts 下的真实字体表' });

const APP_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const DEFAULT_FIXTURES = ['spike-baseline-01', 'spike-baseline-02', 'spike-baseline-03'];
/** DoD 阈值（pt）。L3 断言要求 0.5pt，穿刺阶段用同一把尺子 */
const TOLERANCE_PT = 0.5;
/** 最优假设至少要比次优好这么多倍，否则算「样本分不开」而不是「测出来了」 */
const MIN_MARGIN = 3;

type Hypothesis = 'below' | 'half' | 'prop' | 'above';
const HYPOTHESES: Hypothesis[] = ['below', 'half', 'prop', 'above'];

/**
 * 首行**不是**版心第一行的页 —— 它前面还有一个空段落，量出来的不是基线偏移。
 *
 * 这些页得单独算（见文件末尾的第二段报告）：空段落在 PDF 里**一个片段都不留**，
 * 所以「那一页有没有空段落、它的段落标记挂着什么字体」在真值里查不到，只能照 spec 声明。
 * 这是这类实验的固有代价，不是偷懒：换句话说空段落的行高只能**反推**——
 * 用后一段的基线减去后一段自己的基线偏移。
 */
interface EmptyLead {
  page: number;
  /** 空段落的段落标记：ascii 字体 + 字号（pt）。行高按拉丁规则算就该等于它 */
  markLatin: string;
  markSizePt: number;
  /** 同一个标记的 eastAsia 字体 —— 用来算「若按东亚规则」的对照值 */
  markEastAsia: string;
}

const EMPTY_LEADS: Record<string, EmptyLead[]> = {
  'spike-baseline-01': [
    { page: 8, markLatin: 'TimesNewRomanPSMT', markSizePt: 12, markEastAsia: 'SimSun' },
    { page: 9, markLatin: 'TimesNewRomanPSMT', markSizePt: 22, markEastAsia: 'SimHei' },
  ],
};

/** 一行里的一款字体，单位 twips */
interface Part {
  ascent: number;
  descent: number;
  /** 额外行距：东亚是 win 跨度的 30%，拉丁是 GDI 外部行距 */
  gap: number;
  natural: number;
}

function partsOf(page: TruthPage, lineIndex: number): Part[] {
  const line = page.lines[lineIndex];
  if (!line) return [];
  const eastAsianLine = hasEastAsianText(line.text);
  const seen = new Set<string>();
  const all: { part: Part; eastAsian: boolean }[] = [];

  for (const i of line.items) {
    const item = page.items[i];
    if (!item) continue;
    // 字号取**标称值**（PDF 里 16pt 写成 15.96 是文本矩阵的量化，不是字号真的变了）。
    // 量基线偏移必须用标称值，否则凭空多出 0.3% 的误差；Phase 0 量的是基线**间距**，
    // 两端同误差会抵消，所以那个脚本用 PDF 值反而更稳 —— 两处不一样是故意的。
    const sizePt = Math.round(item.size * 2) / 2;
    const key = `${item.font}@${sizePt}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const { metrics, eastAsian } = loadSpikeFont(item.font);
    const lm = lineMetrics(metrics, ptToTwips(sizePt), { eastAsian: eastAsianLine });
    all.push({
      part: { ascent: lm.ascent, descent: lm.descent, gap: lm.lineGap, natural: lm.lineHeight },
      eastAsian,
    });
  }

  // 东亚行的行盒只由东亚字体决定，拉丁 run 不参与 —— 实测，见 @uw/layout 的 line-height.ts
  const box = eastAsianLine ? all.filter((p) => p.eastAsian) : all;
  return (box.length > 0 ? box : all).map((p) => p.part);
}

/** 某个假设下这一款字体在给定行高里的基线位置 */
function baselineOf(part: Part, lineHeight: number, h: Hypothesis): number {
  const share =
    h === 'below' ? 0 : h === 'above' ? 1 : h === 'half' ? 0.5 : part.ascent / (part.ascent + part.descent);
  const coreAbove = part.ascent + part.gap * share;
  // 行高被网格 / 行距倍数拉大之后多出来的那部分**上下均分** —— 这一条对四个假设都一样，
  // 它们的分歧只在「自然行高里那份额外行距怎么分」。分开写，实验才只测一个变量。
  const core = coreAbove + part.descent + part.gap * (1 - share);
  return coreAbove + (lineHeight - core) / 2;
}

type Group = 'eastAsia' | 'latin';

interface Sample {
  fixture: string;
  page: number;
  label: string;
  group: Group;
  /** 实测：首行基线 − 版心顶 */
  measured: number;
  /** 最终行高：多行段落取实测基线间距，单行段落取自然行高 */
  height: number;
  natural: number;
  predicted: Record<Hypothesis, number>;
}

/** 一行的「字体 + 字号」签名。跨签名的基线差里混着别的段落的行高，不能当行高用 */
function signature(page: TruthPage, lineIndex: number): string {
  const line = page.lines[lineIndex];
  if (!line) return '';
  const parts = new Set<string>();
  for (const i of line.items) {
    const it = page.items[i];
    if (it) parts.add(`${it.font}@${it.size}`);
  }
  return [...parts].sort().join('+');
}

/**
 * 首行的实际行高 —— 取**同签名相邻行**基线差的中位数。
 *
 * 两个限制都是踩过的坑：
 * ① 只在同签名的相邻行之间量。01 的每一页里都换过字体（第二段刻意换成宋体来验证行盒
 *    首尾相接地摞），跨段落的基线差是「上一行的下半 + 这一行的上半」，不是任何一行的行高；
 *    拿它当行高，仿宋 16pt 那页会算出 19.56pt（真值是 20.8pt），整个实验就废了。
 * ② 取中位数而不是平均。开网格 + 1.5 倍行距时 Word 会在 47.76 / 47.64 之间来回摆
 *    （行高 47.7pt 不是 PDF 坐标量化步长的整数倍，靠交替取整保住累计位置），
 *    平均值会被样本个数的奇偶带跑。
 */
function measuredPitch(page: TruthPage): number | undefined {
  const sig = signature(page, 0);
  const deltas: number[] = [];
  for (let i = 1; i < page.lines.length; i++) {
    const prev = page.lines[i - 1];
    const cur = page.lines[i];
    if (!prev || !cur) continue;
    if (signature(page, i) !== sig || signature(page, i - 1) !== sig) continue;
    deltas.push(cur.y - prev.y);
  }
  if (deltas.length === 0) return undefined;
  deltas.sort((a, b) => a - b);
  return deltas[deltas.length >> 1];
}

async function collect(fixture: string): Promise<Sample[]> {
  const truth = JSON.parse(
    await readFile(path.join(APP_ROOT, 'fixtures', `${fixture}.truth.json`), 'utf8'),
  ) as WordTruth;
  const section = truth.sections?.[0];
  if (!section) throw new Error(`${fixture}: 真值里没有 Word 自述的页面设置，取不到版心顶`);

  const skip = new Set((EMPTY_LEADS[fixture] ?? []).map((e) => e.page));
  const samples: Sample[] = [];
  for (const page of truth.pages) {
    if (skip.has(page.index)) continue;
    const first = page.lines[0];
    if (!first) continue;
    const parts = partsOf(page, 0);
    if (parts.length === 0) continue;

    const natural = Math.max(...parts.map((p) => p.natural));
    const pitch = measuredPitch(page);
    // 单行段落没有基线间距可量，只能取自然行高 —— 01 / 02 不开网格、单倍行距，
    // 自然行高就是最终行高，而这一步是 Phase 0 已经验证过的，不是新假设。
    const height = pitch === undefined ? natural : ptToTwips(pitch);

    const predicted = {} as Record<Hypothesis, number>;
    for (const h of HYPOTHESES) {
      predicted[h] = twipsToPt(Math.max(...parts.map((p) => baselineOf(p, height, h))));
    }
    const fonts = [...new Set(first.items.map((i) => `${page.items[i]?.font}@${page.items[i]?.size}`))];
    samples.push({
      fixture,
      page: page.index,
      label: fonts.join('+'),
      group: hasEastAsianText(first.text) ? 'eastAsia' : 'latin',
      measured: first.y - section.topMargin,
      height,
      natural,
      predicted,
    });
  }
  return samples;
}

const names = process.argv.slice(2).filter((a) => !a.startsWith('--'));
const fixtures = names.length > 0 ? names : DEFAULT_FIXTURES;

const samples: Sample[] = [];
for (const f of fixtures) samples.push(...(await collect(f)));

if (samples.length === 0) {
  console.error('没有可比对的样本 —— fixture 里每页都没有文字？');
  process.exit(1);
}

const LABEL_W = 34;
const RULE = '-'.repeat(LABEL_W + 24 + 9 * HYPOTHESES.length);
/** 每组里应该赢的那个假设 —— 与 `@uw/fonts` 的 `baselineOffset` 实现一致，不一致就报错 */
const EXPECTED: Record<Group, Hypothesis> = { eastAsia: 'half', latin: 'above' };
const GROUP_NAME: Record<Group, string> = { eastAsia: '东亚行', latin: '拉丁行' };

let failed = false;

for (const group of ['eastAsia', 'latin'] as Group[]) {
  const rows = samples.filter((s) => s.group === group);
  if (rows.length === 0) continue;

  console.log(`\n${GROUP_NAME[group]} · 首行基线到版心顶（pt），Δ = 实测 − 预测\n`);
  console.log(
    '字体 @ 字号'.padEnd(LABEL_W) +
      '自然'.padStart(8) +
      '行高'.padStart(8) +
      '实测'.padStart(8) +
      HYPOTHESES.map((h) => h.padStart(9)).join(''),
  );
  console.log(RULE);

  const worst: Record<Hypothesis, number> = { below: 0, half: 0, prop: 0, above: 0 };
  let worstSample: Sample | undefined;
  for (const s of rows) {
    const errs = HYPOTHESES.map((h) => s.measured - s.predicted[h]);
    HYPOTHESES.forEach((h, k) => {
      const e = errs[k] as number;
      if (Math.abs(e) > Math.abs(worst[h])) {
        worst[h] = e;
        if (h === EXPECTED[group]) worstSample = s;
      }
    });
    console.log(
      s.label.padEnd(LABEL_W) +
        twipsToPt(s.natural).toFixed(2).padStart(8) +
        twipsToPt(s.height).toFixed(2).padStart(8) +
        s.measured.toFixed(2).padStart(8) +
        errs.map((e) => e.toFixed(3).padStart(9)).join(''),
    );
  }
  console.log(RULE);
  console.log(
    `${rows.length} 个样本，各假设的最大误差：` +
      HYPOTHESES.map((h) => `${h} ${worst[h].toFixed(3)}`).join(' / '),
  );

  const best = HYPOTHESES.reduce((a, b) => (Math.abs(worst[a]) <= Math.abs(worst[b]) ? a : b));
  const runnerUp = HYPOTHESES.filter((h) => h !== best).reduce((a, b) =>
    Math.abs(worst[a]) <= Math.abs(worst[b]) ? a : b,
  );
  const margin = Math.abs(worst[runnerUp]) / Math.abs(worst[best]);
  console.log(`最优假设：${best}，比次优（${runnerUp}）好 ${margin.toFixed(1)} 倍`);
  if (worstSample) {
    console.log(`最差样本：${worstSample.fixture} 第 ${worstSample.page + 1} 页 · ${worstSample.label}`);
  }

  if (best !== EXPECTED[group]) {
    console.error(`✗ ${GROUP_NAME[group]}：代码按 ${EXPECTED[group]} 实现，实测最优却是 ${best}`);
    failed = true;
  } else if (Math.abs(worst[best]) > TOLERANCE_PT) {
    console.error(`✗ ${GROUP_NAME[group]}：最大误差 ${worst[best].toFixed(3)}pt 超过阈值 ${TOLERANCE_PT}pt`);
    failed = true;
  } else if (margin < MIN_MARGIN) {
    console.error(`✗ ${GROUP_NAME[group]}：次优只差 ${margin.toFixed(1)} 倍，样本分不开 —— 加大字号再测`);
    failed = true;
  }
}

// ── 空段落的行高 ──────────────────────────────────────────────────────────────
// 空段落在 PDF 里一个片段都不留，只能用**后一段的基线**减去后一段自己的基线偏移来反推。
// 这一问的答案决定 layout 的 markMetrics() 该取哪个字体桶，而它与直觉相反。
for (const fixture of fixtures) {
  const leads = EMPTY_LEADS[fixture];
  if (!leads) continue;
  const truth = JSON.parse(
    await readFile(path.join(APP_ROOT, 'fixtures', `${fixture}.truth.json`), 'utf8'),
  ) as WordTruth;
  const top = truth.sections?.[0]?.topMargin;
  if (top === undefined) continue;

  console.log(`\n空段落的行高 · ${fixture}（反推：后一段基线 − 后一段自己的基线偏移）\n`);
  console.log(
    '段落标记（ascii/eastAsia@字号）'.padEnd(34) +
      '实测'.padStart(9) +
      '按拉丁'.padStart(9) +
      '按东亚'.padStart(9),
  );
  console.log('-'.repeat(61));

  for (const lead of leads) {
    const page = truth.pages.find((p) => p.index === lead.page);
    const first = page?.lines[0];
    if (!page || !first) continue;
    const parts = partsOf(page, 0);
    if (parts.length === 0) continue;
    const nextNatural = Math.max(...parts.map((p) => p.natural));
    const nextBaseline = twipsToPt(Math.max(...parts.map((p) => baselineOf(p, nextNatural, 'half'))));
    const measured = first.y - top - nextBaseline;

    const size = ptToTwips(lead.markSizePt);
    const asLatin = twipsToPt(
      lineMetrics(loadSpikeFont(lead.markLatin).metrics, size, { eastAsian: false }).lineHeight,
    );
    const asEastAsia = twipsToPt(
      lineMetrics(loadSpikeFont(lead.markEastAsia).metrics, size, { eastAsian: true }).lineHeight,
    );
    console.log(
      `${lead.markLatin}/${lead.markEastAsia}@${lead.markSizePt}`.padEnd(34) +
        measured.toFixed(2).padStart(9) +
        asLatin.toFixed(2).padStart(9) +
        asEastAsia.toFixed(2).padStart(9),
    );
    if (Math.abs(measured - asLatin) > TOLERANCE_PT) {
      console.error(
        `✗ 空段落：实测 ${measured.toFixed(2)}pt 与「按 ascii 桶 + 拉丁规则」的 ${asLatin.toFixed(2)}pt 差太多`,
      );
      failed = true;
    }
    if (Math.abs(measured - asEastAsia) < Math.abs(measured - asLatin)) {
      console.error('✗ 空段落：东亚规则反而更接近 —— layout 的 markMetrics() 该换回 eastAsia 桶');
      failed = true;
    }
  }
}

if (failed) process.exit(1);
console.log(
  [
    '',
    '✓ 基线穿刺通过：',
    '  · 东亚行那 30% 额外行距**上下均分**',
    '  · 拉丁行的 GDI 外部行距**整块在基线以上**',
    '  · 网格吸附与行距倍数拉出来的余量同样上下均分（「行高」列不等于「自然」列的那几行）',
    '  · 空段落走段落标记的 ascii 桶 + 拉丁规则',
    '  行盒可以装了。',
  ].join('\n'),
);
