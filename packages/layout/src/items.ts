/**
 * 段落 → item 流：流水线里「分桶 + 度量」那两步的落地。
 *
 * 一个码点一个 item，不做「一个 run 一个字符串」的合并 —— 逐字 x 是中文排版的硬需求
 * （两端对齐、标点挤压、中西文间距都要逐字微调），而合并回片段是渲染出口的事
 * （见 paragraph.ts 的 fragments）。
 *
 * 这里**不认识 DOM，也不 import 任何字体实现**（原则 1.2）：字体名从 `@uw/fonts` 的
 * `splitFontRuns` 来，宽度从注入的 `TextMeasurer` 来。
 */
import type { Twips } from '@uw/core';
import type { TextMeasurer } from '@uw/fonts';
import { splitFontRuns } from '@uw/fonts';
import type { ResolvedParagraph, ResolvedRun, ResolvedRunProps } from '@uw/model';
import type { KinsokuSets } from './break-class.ts';
import { isCompressiblePunct, isSpaceCp, kinsokuOf } from './break-class.ts';
import type { CharItem, LayoutItem } from './types.ts';
import { AUTO_SPACE_EM, em, SMALL_CAPS_SCALE, VERT_ALIGN_SCALE } from './uncalibrated.ts';

export interface BuildItemsOptions {
  measurer: TextMeasurer;
  /** 禁则集，缺省用内建表。文档自定义的走 `kinsokuFrom(settings)` */
  kinsoku?: KinsokuSets;
  /**
   * 四个字体桶全空时用哪款字体。
   * fonts 包刻意不替调用方决定默认字体（见 `bucketFont`），这个决定在这里做。
   */
  defaultFont?: string;
}

/**
 * 段落 → item 流。
 *
 * 跳过三类内容：隐藏文字（`w:vanish` **不参与排版**，不是画成透明）、域代码
 * （`w:instrText` 是给求值用的，不显示）、域界桩（`w:fldChar` 只标位置）。
 */
export function buildItems(p: ResolvedParagraph, opts: BuildItemsOptions): LayoutItem[] {
  const out: LayoutItem[] = [];
  for (const run of p.runs) {
    if (run.props.hidden) continue;
    appendRun(out, run, opts);
  }
  applyAutoSpace(out, p.props.autoSpaceDE, p.props.autoSpaceDN);
  return out;
}

function appendRun(out: LayoutItem[], run: ResolvedRun, opts: BuildItemsOptions): void {
  const props = run.props;
  const size = effectiveSize(props);
  for (let ci = 0; ci < run.content.length; ci++) {
    const c = run.content[ci] as (typeof run.content)[number];
    switch (c.kind) {
      case 'text':
        appendText(out, run, ci, transformCase(c.text, props), c.text, size, opts);
        break;
      case 'tab':
        out.push({ kind: 'tab', runId: run.id, contentIndex: ci, fontSize: size });
        break;
      case 'break':
        out.push({ kind: 'break', runId: run.id, contentIndex: ci, breakType: c.breakType });
        break;
      case 'symbol': {
        // w:sym 的字体是**片段自己的**，覆盖 run 的四个桶；码位落在 U+F020–U+F0FF 私用区
        // （symbol-encoded 字体的 (3,0) cmap）。按原样查即可 —— 要不要减 0xF000 是
        // FontSource 那一层的事，布局层猜这个只会猜错
        const cp = c.char.codePointAt(0);
        if (cp === undefined) break;
        out.push(single(run, ci, cp, c.font, size, opts));
        break;
      }
      case 'noBreakHyphen': {
        const item = single(run, ci, 0x2011, fontFor(props, 0x2011, opts), size, opts);
        item.noBreak = true;
        out.push(item);
        break;
      }
      case 'softHyphen': {
        // 平时宽度为 0：软连字符只有在此处断行时才显出来，参与排版的是那个「可断」的性质
        const item = single(run, ci, 0x00ad, fontFor(props, 0x2d, opts), size, opts);
        item.softHyphen = true;
        item.width = 0;
        out.push(item);
        break;
      }
      case 'object':
        out.push({
          kind: 'object',
          runId: run.id,
          contentIndex: ci,
          width: c.width,
          height: c.height,
          gapBefore: 0,
        });
        break;
      // fieldChar / fieldInstruction 不占宽度
      default:
        break;
    }
  }
}

function appendText(
  out: LayoutItem[],
  run: ResolvedRun,
  contentIndex: number,
  text: string,
  original: string,
  size: Twips,
  opts: BuildItemsOptions,
): void {
  if (text === '') return;
  const props = run.props;
  const defaultFont = opts.defaultFont ?? '';
  // 小型大写里同一段文字有两个字号（原本的小写字母用缩小的大写字形），批量度量的
  // 前提「一段一个字号」不成立，只能逐字问。它在公文语料里几乎不出现，慢一点无所谓
  const smallCaps = props.smallCaps && !props.caps;
  // 先切成「同字体 + 同脚本」的段，再逐段批量量宽 —— 度量器的热路径是数组进数组出，
  // 逐字调用会把两级缓存的收益吃掉一大半
  for (const fr of splitFontRuns(text, props.fonts)) {
    const slice = text.slice(fr.start, fr.end);
    const cps: number[] = [];
    const offsets: number[] = [];
    const sizes: Twips[] = [];
    let i = fr.start;
    for (const ch of slice) {
      cps.push(ch.codePointAt(0) as number);
      offsets.push(i);
      sizes.push(smallCaps && wasLower(original, i) ? size * SMALL_CAPS_SCALE : size);
      i += ch.length;
    }
    const font = fr.font === '' ? defaultFont : fr.font;
    const widths = new Float64Array(cps.length);
    if (smallCaps) {
      for (let k = 0; k < cps.length; k++) {
        widths[k] = opts.measurer.advance(font, sizes[k] as Twips, cps[k] as number);
      }
    } else {
      opts.measurer.advances(font, size, Uint32Array.from(cps), widths);
    }
    for (let k = 0; k < cps.length; k++) {
      out.push(
        charItem(
          run,
          contentIndex,
          offsets[k] as number,
          cps[k] as number,
          font,
          sizes[k] as Twips,
          fr.script,
          widths[k] as number,
          opts.kinsoku,
        ),
      );
    }
  }
}

/** 大小写变换保证码点数不变（见 transformCase），所以下标可以直接拿回原文查 */
function wasLower(original: string, index: number): boolean {
  const ch = original.slice(index, index + 2).codePointAt(0);
  if (ch === undefined) return false;
  const s = String.fromCodePoint(ch);
  return s !== s.toUpperCase();
}

/** `advance` 传的是**未经 run 属性调整**的字形推进宽度，缩放与字间距在这里统一叠 */
function charItem(
  run: ResolvedRun,
  contentIndex: number,
  offset: number,
  cp: number,
  font: string,
  fontSize: Twips,
  script: CharItem['script'],
  advance: Twips,
  kinsoku: KinsokuSets | undefined,
): CharItem {
  return {
    kind: 'char',
    runId: run.id,
    contentIndex,
    offset,
    cp,
    font,
    fontSize,
    script,
    width: scaledWidth(advance, run.props),
    gapBefore: 0,
    space: isSpaceCp(cp),
    kinsoku: kinsokuOf(cp, kinsoku),
    compressible: isCompressiblePunct(cp),
  };
}

/** 零散的单个字符（符号、连字符）：不走 `splitFontRuns`，单独问一次度量器 */
function single(
  run: ResolvedRun,
  contentIndex: number,
  cp: number,
  font: string,
  size: Twips,
  opts: BuildItemsOptions,
): CharItem {
  const f = font === '' ? (opts.defaultFont ?? '') : font;
  return charItem(
    run,
    contentIndex,
    0,
    cp,
    f,
    size,
    'latin',
    opts.measurer.advance(f, size, cp),
    opts.kinsoku,
  );
}

/**
 * `w:w`（横向缩放，百分比）乘在字形宽度上，`w:spacing`（字间距）是**之后**再加的常量。
 * 顺序反了会让「缩放 50% + 加宽 1pt」的文字宽度差出一截。
 */
function scaledWidth(advance: Twips, props: ResolvedRunProps): Twips {
  return (advance * props.scale) / 100 + props.charSpacing;
}

/** 上下标与小型大写都是**换个字号去量**，两个系数都还没标定，见 uncalibrated.ts */
function effectiveSize(props: ResolvedRunProps): Twips {
  if (props.vertAlign !== 'baseline') return props.size * VERT_ALIGN_SCALE;
  return props.size;
}

/**
 * `w:caps` 把文字整个当大写排 —— 这**改变宽度**，不是渲染层的样式问题。
 *
 * `w:smallCaps` 也走同一个大写化，但**只有原本是小写的那些字符**用缩小的字号
 * （见 appendText 与 SMALL_CAPS_SCALE），这才是「小型大写」而不是「整体缩小」。
 */
function transformCase(text: string, props: ResolvedRunProps): string {
  if (!props.caps && !props.smallCaps) return text;
  const upper = text.toUpperCase();
  // ß → SS 这类一对多的大小写映射会让码点数变化，offset 就对不上原文了。
  // 宁可按原文排（宽度略有偏差），也不能让命中测试与编辑定位错位
  return [...upper].length === [...text].length ? upper : text;
}

/**
 * 该码点在这个 run 里落到哪款字体 —— 只有 `w:noBreakHyphen` / `w:softHyphen` 这种
 * 「没有源文本却要占位」的片段需要单独问一次，正文走 `splitFontRuns` 批量分。
 */
function fontFor(props: ResolvedRunProps, cp: number, opts: BuildItemsOptions): string {
  const runs = splitFontRuns(String.fromCodePoint(cp), props.fonts);
  const font = runs[0]?.font ?? '';
  return font === '' ? (opts.defaultFont ?? '') : font;
}

/**
 * 中西文自动间距（`w:autoSpaceDE` / `w:autoSpaceDN`，两者默认**开**）。
 *
 * 东亚字符与拉丁字母 / 数字相邻时插入 1/8 em。不做的话中英混排的行长永远对不上，
 * 断行点会随着每行的中英切换次数越差越多。
 *
 * 三条边界：
 * - 空格两侧不加 —— 已经有空隙了，再加就成了双份
 * - DE 管字母、DN 管数字，两个开关是分开的（Word 界面上也是两项）
 * - 间距记在**后一个** item 上（`gapBefore`），行首那一个不生效 ——
 *   断行把它俩分到两行时，这个间距必须消失
 */
function applyAutoSpace(items: LayoutItem[], de: boolean, dn: boolean): void {
  if (!de && !dn) return;
  for (let i = 1; i < items.length; i++) {
    const prev = items[i - 1] as LayoutItem;
    const cur = items[i] as LayoutItem;
    if (prev.kind !== 'char' || cur.kind !== 'char') continue;
    if (prev.space || cur.space) continue;

    const eastAsiaSide = prev.script === 'eastAsia' ? prev : cur.script === 'eastAsia' ? cur : undefined;
    const latinSide = prev.script === 'eastAsia' ? cur : prev;
    if (eastAsiaSide === undefined || latinSide.script !== 'latin') continue;

    const digit = latinSide.cp >= 0x30 && latinSide.cp <= 0x39;
    const letter = isLatinLetter(latinSide.cp);
    if (digit ? !dn : letter ? !de : true) continue;

    cur.gapBefore = em(eastAsiaSide.fontSize, AUTO_SPACE_EM);
  }
}

function isLatinLetter(cp: number): boolean {
  return (cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a) || (cp >= 0xc0 && cp <= 0x24f);
}
