import type { RunProps } from './props.ts';

/**
 * 回到段落样式：连字符样式一起移除，语言标记仍跟着内容走。
 * 列全已识别的属性，才能覆盖折叠光标后输入所继承的格式；类型检查保证新增属性不会漏清。
 * 不认识的 XML 仍由回写层保留，不能为了清格式把未知内容一起丢掉。
 */
export const CLEAR_RUN_PROPS = Object.freeze({
  styleId: null,
  fonts: null,
  fontThemes: null,
  bold: null,
  boldCs: null,
  italic: null,
  italicCs: null,
  caps: null,
  smallCaps: null,
  strike: null,
  doubleStrike: null,
  hidden: null,
  size: null,
  sizeCs: null,
  underline: null,
  color: null,
  themeColor: null,
  vertAlign: null,
  charSpacing: null,
  scale: null,
  position: null,
  kerning: null,
  snapToGrid: null,
} satisfies Record<Exclude<keyof RunProps, 'langEastAsia'>, null>);
