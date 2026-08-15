/**
 * OOXML 取值的小助手。
 *
 * 单独一个文件是因为**布尔的规则反直觉**，而它遍布整份文档：写错一次，
 * `<w:b/>`（加粗）与 `<w:b w:val="0"/>`（明确不加粗）会被当成同一件事。
 */
import type { XmlElement } from '@uw/ooxml';
import { attr, child } from '@uw/ooxml';

/**
 * `ST_OnOff`（ECMA-376 §17.17.4）：
 * - 元素在但**没有** `w:val` → **true**（`<w:b/>` 就是加粗）
 * - `w:val` 是 `0` / `false` / `off` → false
 * - 其余 → true
 *
 * 注意「元素不在」和「元素在且为 false」是两回事：前者是「这层没意见，继承上层」，
 * 后者是「这层明确关掉」。所以返回 undefined 而不是 false。
 */
export function onOff(parent: XmlElement, name: string): boolean | undefined {
  const el = child(parent, name);
  if (el === undefined) return undefined;
  const v = attr(el, 'w:val');
  if (v === undefined) return true;
  return v !== '0' && v !== 'false' && v !== 'off';
}

/** `<w:x w:val="..."/>` 的字符串值 */
export function valOf(parent: XmlElement, name: string): string | undefined {
  const el = child(parent, name);
  return el === undefined ? undefined : attr(el, 'w:val');
}

/** `<w:x w:val="123"/>` 的整数值。非法数字按「没写」处理，不抛 */
export function intVal(parent: XmlElement, name: string): number | undefined {
  return toInt(valOf(parent, name));
}

/** 容忍元素缺席的属性取值 —— `w:pgSz` 这类「整个元素可能不在」的场合到处都是 */
export function attrOf(el: XmlElement | undefined, name: string): string | undefined {
  return el === undefined ? undefined : attr(el, name);
}

export function attrInt(el: XmlElement | undefined, name: string): number | undefined {
  return toInt(attrOf(el, name));
}

export function toInt(v: string | undefined): number | undefined {
  if (v === undefined) return undefined;
  const n = Number.parseInt(v, 10);
  return Number.isNaN(n) ? undefined : n;
}

/** 属性版的 ST_OnOff（`w:beforeAutospacing="1"` 这类挂在属性上的开关） */
export function attrOnOff(el: XmlElement, name: string): boolean | undefined {
  const v = attr(el, name);
  if (v === undefined) return undefined;
  return v !== '0' && v !== 'false' && v !== 'off';
}

/**
 * 只在值有定义时才写进对象。
 *
 * `exactOptionalPropertyTypes` 开着，`o.k = undefined` 是类型错误；
 * 而「显式写 undefined」和「不写」在级联里语义不同（见 onOff 的说明），
 * 所以必须是「不写」。
 */
export function put<T extends object, K extends keyof T>(o: T, k: K, v: T[K] | undefined): void {
  if (v !== undefined) o[k] = v;
}

/** 值在白名单里才返回，否则当没写 —— 认不出的枚举值是内容问题，不该让整篇文档挂掉 */
export function enumVal<T extends string>(v: string | undefined, allowed: readonly T[]): T | undefined {
  return v !== undefined && (allowed as readonly string[]).includes(v) ? (v as T) : undefined;
}
