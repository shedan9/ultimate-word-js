/**
 * 段落边框（`w:pBdr`）与段落底纹（`w:pPr/w:shd`）的几何。
 *
 * 分两步，因为框**看邻居**而段落布局要能缓存：
 * 1. `standaloneFrame()` —— 在 `layoutParagraph` 里给出「单独一段」的框：上下边都画、
 *    都让出高度。与邻居无关，缓存命中时原样拿出来
 * 2. `joinParagraphFrames()` —— 摞块的那一方（分页、单元格、页眉页脚）按相邻段落成组：
 *    组内不画各自的上下边（有 `between` 时画它），让出的高度跟着改，底纹与竖线接过段间空当
 *
 * 让出的高度直接折进 `spaceBefore` / `spaceAfter`（见 `ParagraphLayout` 的注释），
 * 于是凡是「把块往下摞」的地方都不用认识边框；只有分页要把两者拆开 —— 段前间距在页首不算，
 * 边框让出的高度照算。
 *
 * **这一层没有真值**（本机没有 Word），几条规则都是照 Word 界面行为与规范写的，
 * 集中列在 `uncalibrated.ts` 的「段落边框」一节，每条写了拿什么样本能钉死。
 */
import type { Twips } from '@uw/core';
import type { Border, ParagraphBorders, ResolvedParaProps } from '@uw/model';
import type { ParagraphFrame, ParagraphLayout } from './types.ts';
import { borderThicknessFactor, PARA_FRAME_GROUP_KEYS } from './uncalibrated.ts';

const SIDES = ['top', 'left', 'bottom', 'right', 'between'] as const;

/** `nil` / `none` 是「明确没有」：它们在级联里盖掉样式给的线，到这儿就该消失 */
function visible(b: Border | undefined): Border | undefined {
  if (b === undefined || b.style === 'nil' || b.style === 'none') return undefined;
  return b;
}

/** 一条线让出的高度：画出来的厚度（双线是 3 × `w:sz`，与表格格线同一张表）+ 它与文字的间距 */
export function borderExtent(b: Border | undefined): Twips {
  return b === undefined ? 0 : b.size * borderThicknessFactor(b.style) + b.space;
}

/**
 * 单独一段的框。没有可见的边、也没有底纹时答 undefined —— 绝大多数段落走这条，
 * `ParagraphLayout` 上连这个字段都不出现。
 *
 * `x` / `width` 是文字区：左取首行与其余行左边缘里**更靠左**的那个（悬挂缩进的编号在框里），
 * 右取右缩进。
 */
export function standaloneFrame(
  props: ResolvedParaProps,
  x: Twips,
  width: Twips,
): ParagraphFrame | undefined {
  const borders: ParagraphBorders = {};
  for (const side of SIDES) {
    const b = visible(props.borders[side]);
    if (b !== undefined) borders[side] = b;
  }
  const shading = props.shading === undefined || props.shading.pattern === 'nil' ? undefined : props.shading;
  if (Object.keys(borders).length === 0 && shading === undefined) return undefined;
  const frame: ParagraphFrame = {
    x,
    width,
    borders,
    insetTop: borderExtent(borders.top),
    insetBottom: borderExtent(borders.bottom),
    joinPrev: false,
    joinNext: false,
  };
  if (shading !== undefined) frame.shading = shading;
  if (borders.top !== undefined) frame.top = borders.top;
  if (borders.bottom !== undefined) frame.bottom = borders.bottom;
  return frame;
}

/** 两段能不能框在一起：边框、底纹、左右边都一样（`PARA_FRAME_GROUP_KEYS`） */
function sameGroup(a: ParagraphFrame | undefined, b: ParagraphFrame | undefined): boolean {
  if (a === undefined || b === undefined) return false;
  return PARA_FRAME_GROUP_KEYS.every((k) => JSON.stringify(a[k]) === JSON.stringify(b[k]));
}

/**
 * 按邻居成组。`undefined` 占位的是表格（或别的非段落块）—— 它把组截断。
 *
 * 不改入参：`layoutParagraph` 的结果可能来自缓存，改了会串到别的文档里去。
 * 没变的段落原样返回同一个对象。
 */
export function joinParagraphFrames(
  layouts: readonly (ParagraphLayout | undefined)[],
): (ParagraphLayout | undefined)[] {
  return layouts.map((layout, i) => {
    const frame = layout?.frame;
    if (layout === undefined || frame === undefined) return layout;
    const joinPrev = sameGroup(layouts[i - 1]?.frame, frame);
    const joinNext = sameGroup(frame, layouts[i + 1]?.frame);
    if (joinPrev === frame.joinPrev && joinNext === frame.joinNext) return layout;
    const top = joinPrev ? frame.borders.between : frame.borders.top;
    const bottom = joinNext ? undefined : frame.borders.bottom;
    const next: ParagraphFrame = {
      ...frame,
      insetTop: borderExtent(top),
      insetBottom: borderExtent(bottom),
      joinPrev,
      joinNext,
    };
    delete next.top;
    delete next.bottom;
    if (top !== undefined) next.top = top;
    if (bottom !== undefined) next.bottom = bottom;
    return {
      ...layout,
      frame: next,
      spaceBefore: layout.spaceBefore - frame.insetTop + next.insetTop,
      spaceAfter: layout.spaceAfter - frame.insetBottom + next.insetBottom,
    };
  });
}

/**
 * 段落被切开时切口那一侧的框：不画线、不让高度（跨页 / 拆行的框不封口，见 `uncalibrated.ts`）。
 * 段前 / 段后间距由调用方处理，这里只动框。
 */
export function openFrame(
  frame: ParagraphFrame | undefined,
  side: 'top' | 'bottom',
): ParagraphFrame | undefined {
  if (frame === undefined) return undefined;
  const next: ParagraphFrame = { ...frame };
  if (side === 'top') {
    delete next.top;
    next.insetTop = 0;
  } else {
    delete next.bottom;
    next.insetBottom = 0;
    next.joinNext = false;
  }
  return next;
}
