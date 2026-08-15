/**
 * 行高 —— 只算**总量**，不定基线。
 *
 * 「单倍行距等于多少」在 Phase 0 已经用 13 个样本标定死了（`@uw/fonts` 的 `lineMetrics`），
 * 这里做的是它之上的三件事：按 `w:lineRule` 解释 `w:line`、多字体逐项取 max、按行网格吸附。
 *
 * ⚠️ 基线在这个高度里的位置**仍未标定**（东亚那 30% 额外行距怎么分到基线上下），
 * 所以这里只回答「这一行占多高」，不回答「文字画在哪儿」。行盒装配等那次穿刺。
 */
import type { Twips } from '@uw/core';
import type { LineMetrics, TextMeasurer } from '@uw/fonts';
import { bucketFont, combineLineMetrics } from '@uw/fonts';
import type { DocGrid, ResolvedParaProps } from '@uw/model';
import type { LayoutItem } from './types.ts';

export interface LineHeightContext {
  measurer: TextMeasurer;
  docGrid: DocGrid;
  defaultFont?: string;
}

export interface LineHeight {
  metrics: LineMetrics;
  /** 行距规则与网格吸附之后的最终行高 */
  height: Twips;
}

/**
 * 一行的行高。
 *
 * 「这一行算不算东亚行」按**整行**判定（只要有一个东亚字就算），与 `@uw/fonts` 里
 * `hasEastAsianText` 的注释一致 —— Word 是按行选行距规则的，不是按 run。
 */
export function lineHeight(
  items: readonly LayoutItem[],
  range: { start: number; end: number },
  props: ResolvedParaProps,
  ctx: LineHeightContext,
): LineHeight {
  const eastAsian = hasEastAsia(items, range);
  const parts: LineMetrics[] = [];
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
    parts.push(ctx.measurer.lineMetrics(font, item.fontSize, { eastAsian }));
  }

  if (parts.length === 0) {
    // 空段落的行高全靠段落标记（¶）自己的字符属性，它不是摆设
    parts.push(markMetrics(props, ctx));
  }
  const metrics = combineLineMetrics(parts);

  // 内嵌对象坐在基线上，整个高度都算进 ascent 一侧。这在基线定下来之前只影响行高总量，
  // 所以先这么记着：至少「有图的行更高」这件事是对的
  const natural = Math.max(metrics.lineHeight, objectHeight);
  return { metrics, height: applyGrid(applyLineRule(natural, props), props, ctx.docGrid) };
}

/**
 * `w:lineRule` 三分支。`w:line` 的刻度取决于它，这也是 model 层故意不转这个单位的原因：
 * - `auto`：1/240 行，240 = 单倍
 * - `exact` / `atLeast`：twips
 */
function applyLineRule(natural: Twips, props: ResolvedParaProps): Twips {
  const { line, lineRule } = props.spacing;
  if (lineRule === 'exact') return line;
  if (lineRule === 'atLeast') return Math.max(natural, line);
  return (natural * line) / 240;
}

/**
 * 行网格吸附 —— 中文公文「每页 22 行」的实现，也是 Phase 0 穿刺踩过的坑：
 * 网格一开，基线被吸到 `linePitch` 的整数倍上，字体度量的差异整个被盖掉。
 *
 * 三条边界：
 * - `type="default"` 不吸（那就是「没有网格」）
 * - `w:lineRule="exact"`（固定值行距）不吸：用户既然写死了行高，网格不该再改它
 * - 段落关了 `w:snapToGrid` 不吸
 *
 * ⚠️ 吸附与行距倍数的**先后**没有真值验证：这里是先乘倍数再吸到整数倍。
 * 「1.5 倍行距 + 开网格」的样本能把它钉死，写在这里免得被当成实测结论。
 */
function applyGrid(height: Twips, props: ResolvedParaProps, grid: DocGrid): Twips {
  if (!props.snapToGrid) return height;
  if (grid.type === 'default' || grid.linePitch <= 0) return height;
  if (props.spacing.lineRule === 'exact') return height;
  const n = Math.max(1, Math.ceil(height / grid.linePitch - 1e-6));
  return n * grid.linePitch;
}

function hasEastAsia(items: readonly LayoutItem[], range: { start: number; end: number }): boolean {
  for (let i = range.start; i < range.end; i++) {
    const item = items[i] as LayoutItem;
    if (item.kind === 'char' && item.script === 'eastAsia') return true;
  }
  return false;
}

/**
 * 空行的度量取段落标记的字符属性。
 *
 * 字体优先取东亚桶：中文文档里段落标记几乎总带着东亚字体，用 ascii 桶会让空行
 * 比周围的行矮一截。「算不算东亚行」同理按东亚桶是否有字体判断 —— 这是个判断，
 * 不是实测结论；空段落的行高值得在基线穿刺时顺手一起量。
 */
function markMetrics(props: ResolvedParaProps, ctx: LineHeightContext): LineMetrics {
  const fonts = props.markRunProps.fonts;
  const eastAsia = bucketFont(fonts, 'eastAsia');
  const font = eastAsia !== '' ? eastAsia : bucketFont(fonts, 'ascii');
  return ctx.measurer.lineMetrics(font === '' ? (ctx.defaultFont ?? '') : font, props.markRunProps.size, {
    eastAsian: eastAsia !== '',
  });
}
