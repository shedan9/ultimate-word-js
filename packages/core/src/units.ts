/**
 * 单位系统。
 *
 * 布局全程用 **twips**（1/1440 英寸）—— OOXML 原生就是 twips / half-point / EMU，
 * 用它做累加不会引入换算误差。px 只允许在渲染层最后一次转换时出现；布局过程中
 * 出现 px 一律视为 bug（行高、页高逐行累加时的浮点漂移会让最后一行溢出）。
 */

/** 1/1440 英寸。布局计算的唯一单位 */
export type Twips = number;
/** 1/72 英寸。仅用于与 PDF / Word UI 交互 */
export type Points = number;
/** CSS 像素，仅存在于渲染层 */
export type Pixels = number;

export const TWIP_PER_INCH = 1440;
export const TWIP_PER_PT = 20;
/** OOXML 里字号用半磅（w:sz="32" = 16pt） */
export const TWIP_PER_HALF_PT = 10;
/** EMU（English Metric Unit）：图片尺寸用，1 pt = 12700 EMU */
export const EMU_PER_TWIP = 635;
export const EMU_PER_INCH = 914400;
/** CSS 参考像素固定 96dpi，与设备 DPI 无关 */
export const PX_PER_INCH = 96;
export const TWIP_PER_PX = TWIP_PER_INCH / PX_PER_INCH; // 15
export const TWIP_PER_MM = TWIP_PER_INCH / 25.4;
export const TWIP_PER_CM = TWIP_PER_MM * 10;

export const ptToTwips = (pt: Points): Twips => pt * TWIP_PER_PT;
export const twipsToPt = (tw: Twips): Points => tw / TWIP_PER_PT;
/** w:sz / w:szCs 的半磅值 → twips */
export const halfPtToTwips = (halfPt: number): Twips => halfPt * TWIP_PER_HALF_PT;
export const twipsToHalfPt = (tw: Twips): number => tw / TWIP_PER_HALF_PT;
export const emuToTwips = (emu: number): Twips => emu / EMU_PER_TWIP;
export const twipsToEmu = (tw: Twips): number => tw * EMU_PER_TWIP;
export const inchToTwips = (inch: number): Twips => inch * TWIP_PER_INCH;
export const mmToTwips = (mm: number): Twips => mm * TWIP_PER_MM;
export const cmToTwips = (cm: number): Twips => cm * TWIP_PER_CM;

/** 唯一允许产出 px 的方向：渲染层出口 */
export const twipsToPx = (tw: Twips, zoom = 1): Pixels => (tw / TWIP_PER_PX) * zoom;
/** 反向只用于把鼠标坐标打回模型空间 */
export const pxToTwips = (px: Pixels, zoom = 1): Twips => (px / zoom) * TWIP_PER_PX;

/**
 * 字体设计单位 → twips。
 * `unitsPerEm` 是字体自己的坐标系（TrueType 常见 1000 / 1024 / 2048）。
 */
export const fontUnitsToTwips = (value: number, unitsPerEm: number, fontSize: Twips): Twips =>
  (value / unitsPerEm) * fontSize;
