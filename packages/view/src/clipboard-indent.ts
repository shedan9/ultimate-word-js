import { TWIP_PER_PT } from '@uw/core';
import type { Indent } from '@uw/model';

/** 显式零要盖住目标样式中的字符单位与悬挂；删除直接格式会重新露出目标样式。 */
export const ZERO_INDENT: Readonly<Required<Indent>> = {
  left: 0,
  right: 0,
  firstLine: 0,
  hanging: 0,
  leftChars: 0,
  rightChars: 0,
  firstLineChars: 0,
  hangingChars: 0,
};

/** CSS 长度只接受绝对单位与无单位零，不把百分比或容器宽度误当成文档缩进。 */
function length(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const match = /^([+-]?(?:\d+(?:\.\d+)?|\.\d+))(pt|px|in|cm|mm|pc)?$/i.exec(value.trim());
  if (!match?.[1]) return undefined;
  const n = Number(match[1]);
  if (!match[2] && n !== 0) return undefined;
  const unit = match[2]?.toLowerCase() ?? 'pt';
  const factors: Record<string, number> = {
    pt: 20,
    px: 15,
    in: 1440,
    cm: 1440 / 2.54,
    mm: 144 / 2.54,
    pc: 240,
  };
  const twips = Math.round(n * (factors[unit] ?? TWIP_PER_PT));
  return Number.isSafeInteger(twips) ? twips : undefined;
}

function count(value: string | undefined): number | undefined {
  if (value === undefined || !/^[+-]?(?:\d+(?:\.\d+)?|\.\d+)$/.test(value.trim())) return undefined;
  const chars = Math.round(Number(value) * 100);
  return Number.isSafeInteger(chars) ? chars : undefined;
}

/** 只供自己的复制片段保留全部八个字段（包括被字符单位盖住的兼容值），坏标记退回 CSS。 */
function ownIndent(el: Element): Required<Indent> | undefined {
  const raw = el.getAttribute('data-uw-indent');
  if (!raw || !el.closest('[data-uw-fragment]')) return undefined;
  try {
    const value: unknown = JSON.parse(raw);
    if (!value || typeof value !== 'object' || Array.isArray(value)) return undefined;
    const out = { ...ZERO_INDENT };
    for (const key of Object.keys(out) as (keyof Indent)[]) {
      const n = (value as Record<string, unknown>)[key];
      if (typeof n !== 'number' || !Number.isSafeInteger(n)) return undefined;
      out[key] = n;
    }
    return out;
  } catch {
    return undefined;
  }
}

/** 左右边距不继承；只有 text-indent 的首行 / 悬挂部分沿父级继承。行内元素不改段落缩进。 */
export function readClipboardIndent(
  el: Element,
  style: Map<string, string>,
  inherited: Readonly<Required<Indent>> = ZERO_INDENT,
): Required<Indent> {
  const own = ownIndent(el);
  if (own) return own;
  const out = {
    ...ZERO_INDENT,
    firstLine: inherited.firstLine,
    hanging: inherited.hanging,
    firstLineChars: inherited.firstLineChars,
    hangingChars: inherited.hangingChars,
  };
  for (const side of ['left', 'right'] as const) {
    const value = style.get(`mso-para-margin-${side}`) ?? style.get(`margin-${side}`);
    const chars = value?.trim().match(/^([+-]?(?:\d+(?:\.\d+)?|\.\d+))gd$/i);
    const grid = count(chars?.[1]);
    if (grid !== undefined) out[`${side}Chars`] = grid;
    else out[side] = length(value) ?? length(style.get(`margin-${side}`)) ?? 0;
  }
  const first = length(style.get('text-indent'));
  const chars = count(style.get('mso-char-indent-count'));
  if (first !== undefined || chars !== undefined) {
    out.firstLine = first !== undefined && first > 0 ? first : 0;
    out.hanging = first !== undefined && first < 0 ? -first : 0;
    out.firstLineChars = chars !== undefined && chars > 0 ? chars : 0;
    out.hangingChars = chars !== undefined && chars < 0 ? -chars : 0;
    // 字符单位决定方向；旧兼容长度不能把首行缩进误变为悬挂，零同样明确取消继承。
    if (chars !== undefined) {
      if (chars > 0) out.hanging = 0;
      if (chars < 0) out.firstLine = 0;
    }
  }
  return out;
}

/** CSS 保证其他应用能显示缩进，Word 私有属性保留字符语义；精确的兼容字段另由片段标记携带。 */
export function clipboardIndentStyle(indent: Indent): string {
  const ind = { ...ZERO_INDENT, ...indent };
  const offset = (twips: number, chars: number) => (chars ? `${chars / 100}em` : `${twips / TWIP_PER_PT}pt`);
  const hanging = ind.hangingChars !== 0 || ind.hanging !== 0;
  const first = hanging ? -ind.hanging : ind.firstLine;
  const chars = hanging ? -ind.hangingChars : ind.firstLineChars;
  return [
    `margin-left:${offset(ind.left, ind.leftChars)}`,
    `margin-right:${offset(ind.right, ind.rightChars)}`,
    `text-indent:${offset(first, chars)}`,
    ...(ind.leftChars ? [`mso-para-margin-left:${ind.leftChars / 100}gd`] : []),
    ...(ind.rightChars ? [`mso-para-margin-right:${ind.rightChars / 100}gd`] : []),
    ...(chars ? [`mso-char-indent-count:${chars / 100}`] : []),
  ].join(';');
}
