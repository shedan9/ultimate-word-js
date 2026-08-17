/**
 * 行高与基线 —— 一行占多高，以及文字画在这个高度里的哪个位置。
 *
 * 两问都已由穿刺定死，实现分两边：**总量**在 `@uw/fonts` 的 `lineMetrics`（Phase 0，13 个样本），
 * **基线**在同一个包的 `baselineOffset`（基线穿刺，26 个样本）。这里做的是它们之上的四件事：
 * 按 `w:lineRule` 解释 `w:line`、多字体合成、按行网格吸附、把段落标记的度量补上。
 *
 * 顺序是这一层最容易搞反的地方，见 `applyLineRule` 的注释：网格吸附在**行距倍数之前**。
 */
import type { Twips } from '@uw/core';
import type { LineMetrics, TextMeasurer } from '@uw/fonts';
import { bucketFont, composeBaseline, naturalLineHeight } from '@uw/fonts';
import type { DocGrid, ResolvedParaProps } from '@uw/model';
import type { LayoutItem } from './types.ts';

export interface LineHeightContext {
  measurer: TextMeasurer;
  docGrid: DocGrid;
  defaultFont?: string;
}

export interface LineHeight {
  /** 行距规则与网格吸附之后的最终行高 */
  height: Twips;
  /** 行顶到基线，`0 <= baseline <= height` */
  baseline: Twips;
  /** 未经行距规则与网格调整的自然行高，回归比对时要能分辨「是规则的锅还是度量的锅」 */
  natural: Twips;
}

/**
 * 一行的行高与基线。
 *
 * 「这一行算不算东亚行」按**整行**判定（只要有一个东亚字就算），与 `@uw/fonts` 里
 * `hasEastAsianText` 的注释一致 —— Word 是按行选行距规则的，不是按 run。
 *
 * 东亚行的行盒**只由东亚桶的字体决定**，拉丁 run 完全不参与 —— 这一条反直觉，是实测的
 * （spike-baseline-02 末四页，中西混排行的首行基线与「同字号纯东亚」那几页**一模一样**，
 * 小数点后三位都不差）。最能说明问题的是等线 72pt 配 Times New Roman 72pt 那一页：
 * Times 的 winAscent（64.16pt）比等线的（58.32pt）大，若它参与合成就该赢，
 * 实测基线仍落在等线单独算出来的 69.57pt 上。
 *
 * 这也解释了 Phase 0 的那条结论「含东亚文字的行不加外部行距」：不是「外部行距被扣掉」，
 * 而是拉丁字体压根没进这一行的行盒。
 */
export function lineHeight(
  items: readonly LayoutItem[],
  range: { start: number; end: number },
  props: ResolvedParaProps,
  ctx: LineHeightContext,
): LineHeight {
  const eastAsian = hasEastAsia(items, range);
  /** 定行盒的那些字体：东亚行里只有东亚桶，拉丁行里就是全部 */
  const box: LineMetrics[] = [];
  /** 不定行盒但也不能被切掉的那些（东亚行里的拉丁 run），只用来兜底 */
  const passenger: LineMetrics[] = [];
  const seen = new Set<string>();
  let objectHeight = 0;

  for (let i = range.start; i < range.end; i++) {
    const item = items[i] as LayoutItem;
    if (item.kind === 'object') {
      if (item.height > objectHeight) objectHeight = item.height;
      continue;
    }
    if (item.kind === 'break') continue;
    const font = item.kind === 'char' ? item.font : '';
    const key = `${font}|${item.fontSize}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const m = ctx.measurer.lineMetrics(font, item.fontSize, { eastAsian });
    const decidesBox = !eastAsian || (item.kind === 'char' && item.script === 'eastAsia');
    (decidesBox ? box : passenger).push(m);
  }

  if (box.length === 0) {
    // 空段落的行高全靠段落标记（¶）自己的字符属性，它不是摆设。
    // 「东亚行但一个东亚 item 都没有」不可能（eastAsian 就是这么判的），所以这里只会是空行。
    box.push(markMetrics(props, ctx));
  }

  // 内嵌对象坐在基线上，整个高度都算进基线以上 —— 所以它进「行至少这么高」与
  // 「基线至少这么低」两处，但**不参与居中**：核心盒居中那条规则是给文字量的，
  // 把图片也居中会让它凭空浮起来。
  // ⚠️ 未标定：一份「一行里一张 20pt 高的图 + 五号字」的样本能钉死它，优先级低于文字。
  const floor = floorBox(passenger);
  const natural = Math.max(naturalLineHeight(box), objectHeight, floor.height);
  const height = applyLineRule(applyGrid(natural, props, ctx.docGrid), props);
  const baseline = Math.min(height, Math.max(composeBaseline(box, height), objectHeight, floor.above));
  return { height, baseline, natural };
}

/**
 * `w:lineRule` 三分支。`w:line` 的刻度取决于它，这也是 model 层故意不转这个单位的原因：
 * - `auto`：1/240 行，240 = 单倍
 * - `exact` / `atLeast`：twips
 *
 * 入参 `height` 是**已经吸附过网格**的高度 —— 顺序是实测的，见 `applyGrid`。
 *
 * ⚠️ `atLeast` 与网格的先后没有真值：这里把 `w:line` 当成吸附**之后**的下限
 * （`max(吸附后, line)`），另一种可能是先取下限再吸附。公文里 `atLeast` 罕见，
 * 一份「atLeast 20pt + 网格 31.8pt」的样本就能钉死，暂不为它设计。
 */
function applyLineRule(height: Twips, props: ResolvedParaProps): Twips {
  const { line, lineRule } = props.spacing;
  if (lineRule === 'exact') return line;
  if (lineRule === 'atLeast') return Math.max(height, line);
  return (height * line) / 240;
}

/**
 * 行网格吸附 —— 中文公文「每页 22 行」的实现，也是 Phase 0 穿刺踩过的坑：
 * 网格一开，基线被吸到 `linePitch` 的整数倍上，字体度量的差异整个被盖掉。
 *
 * 三条边界：
 * - `type="default"` 不吸（那就是「没有网格」）
 * - `w:lineRule="exact"`（固定值行距）不吸：用户既然写死了行高，网格不该再改它
 * - 段落关了 `w:snapToGrid` 不吸（实测：spike-baseline-03 第五段退回自然行高 20.76pt）
 *
 * **吸附在行距倍数之前**，这一条与直觉相反，是实测的（spike-baseline-03 末三段）：
 * 网格 31.8pt 下开 1.5 倍行距，仿宋 16pt 与宋体 12pt 的行高**都是 47.7pt**
 * （= 1.5 × 31.8），与字号无关。若是先乘倍数再吸附，两者分别是 31.2 与 23.4pt，
 * 都会吸到 31.8pt —— 差着半行，而且会随字号变。
 */
function applyGrid(height: Twips, props: ResolvedParaProps, grid: DocGrid): Twips {
  if (!props.snapToGrid) return height;
  if (grid.type === 'default' || grid.linePitch <= 0) return height;
  if (props.spacing.lineRule === 'exact') return height;
  const n = Math.max(1, Math.ceil(height / grid.linePitch - 1e-6));
  return n * grid.linePitch;
}

/**
 * 不定行盒的那些字体所需的最小 ascent —— 一条**防切字**的下限，不是实测规则。
 *
 * 实测样本里它永远不生效：中西混排且同字号时，东亚一侧算出来的基线（≥0.96 em）
 * 总是比拉丁一侧的核心盒上沿（≤0.94 em）低，所以加不加它，26 个样本的预测值一个都不变。
 * 加它是为了「12pt 汉字里嵌一个 72pt 英文单词」这种样本外的情形不至于让字叠在上一行 ——
 * 那时 Word 大概会把行撑高，但我们没有真值，宁可保守地保证不切字。
 *
 * 补一份「东亚小字号 + 拉丁大字号同行」的样本就能把这个 `Math.max` 换成真规则；
 * 在那之前它是**判断**，所以单独一个函数、单独一段注释，不许混进上面的实测公式里。
 */
function floorBox(parts: readonly LineMetrics[]): { above: Twips; height: Twips } {
  let above = 0;
  let height = 0;
  for (const p of parts) {
    if (p.coreAbove > above) above = p.coreAbove;
    // 下沿也要留出来，否则 g / y 的尾巴会被下一行盖掉
    if (p.coreAbove + p.descent > height) height = p.coreAbove + p.descent;
  }
  return { above, height };
}

function hasEastAsia(items: readonly LayoutItem[], range: { start: number; end: number }): boolean {
  for (let i = range.start; i < range.end; i++) {
    const item = items[i] as LayoutItem;
    if (item.kind === 'char' && item.script === 'eastAsia') return true;
  }
  return false;
}

/**
 * 空行的度量取段落标记的字符属性，走 **ascii 桶 + 拉丁规则**。
 *
 * 这是实测结论，且与直觉相反（spike-baseline-01 末两页）：一个只有段落标记的空段落，
 * 标记同时挂着 `w:eastAsia="宋体"` 与 `w:ascii="Times New Roman"`，Word 给的行高是
 * 13.82pt @12pt —— 那是 Times New Roman 的拉丁行高（1.1499 em），不是宋体的 15.6pt
 * （1.3 em）。空的黑体 22pt 段落同理给 25.34pt = Times 22pt，不是 28.6pt。
 *
 * 原因不神秘：段落标记本身不是东亚字符，逐字符分桶把它分到 ascii 桶，
 * 于是「这一行有没有东亚字」的答案是「没有」，`lineMetrics` 也就不该乘 1.3。
 * 换句话说这里不需要特例，需要的是**别自作聪明去优先东亚桶**。
 *
 * ⚠️ 未标定：`w:hint="eastAsia"` 会不会把答案翻过来。COM 不方便直接写 hint，
 * 要手改 XML 造样本；公文里空段落几乎都是这个配置，所以先按实测的来。
 */
function markMetrics(props: ResolvedParaProps, ctx: LineHeightContext): LineMetrics {
  const fonts = props.markRunProps.fonts;
  const ascii = bucketFont(fonts, 'ascii');
  const font = ascii !== '' ? ascii : bucketFont(fonts, 'eastAsia');
  return ctx.measurer.lineMetrics(font === '' ? (ctx.defaultFont ?? '') : font, props.markRunProps.size, {
    eastAsian: false,
  });
}
