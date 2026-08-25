/**
 * Word 真值（truth）数据结构。
 *
 * 全部坐标单位为 **pt（1/72 英寸）**，原点为**页面左上角、y 轴向下** —— 与 Word
 * 的版式直觉一致，而不是 PDF 原生的左下原点。布局引擎内部用 twips，比对层负责
 * 换算（1 pt = 20 twips），本文件不引入 twips 以免真值里出现两套单位。
 */

export interface TruthGenerator {
  /** 生成工具与其格式版本；格式不兼容变更时 +1，比对层据此拒绝旧真值 */
  tool: string;
  formatVersion: number;
  pdfjs: string;
  /** 来自 Word COM 的版本号，如 "16.0"；无 sidecar 时缺省 */
  word?: string;
  wordBuild?: string;
}

/** Word 自己报告的节页面设置（sidecar JSON，非从 PDF 反推） */
export interface WordSection {
  pageWidth: number;
  pageHeight: number;
  topMargin: number;
  bottomMargin: number;
  leftMargin: number;
  rightMargin: number;
  headerDist: number;
  footerDist: number;
  gutter: number;
  orientation: number;
  lineNumbers: boolean;
}

export interface WordMeta {
  source: string;
  wordVersion: string;
  wordBuild: string;
  unit: 'pt';
  pageCount: number;
  wordCount: number;
  charCount: number;
  sections: WordSection[];
}

/**
 * PDF 里实际用到的字体。ascent / descent 是 pdf.js 从字体表读出的、已除以
 * unitsPerEm 的归一化值 —— 这正是 Word 算行高的输入，Phase 0 的行高验证直接拿它对。
 */
export interface TruthFont {
  /** 已剥掉子集前缀（"BCDEEE+"）的字体名；子集前缀每次导出都会变，留着会污染 diff */
  name: string;
  ascent: number;
  descent: number;
  vertical: boolean;
}

/** 一个文本片段：PDF 里一次 show-text 操作的产物，粒度约等于「同字体的连续字符段」 */
export interface TruthItem {
  /** 片段左端 x（pt，页面左上角原点） */
  x: number;
  /** 基线 y（pt，从页顶向下） */
  y: number;
  /** 片段推进宽度（pt） */
  w: number;
  /** 字号（pt），由文本矩阵的纵向缩放推出 */
  size: number;
  /** 字体名，对应 WordTruth.fonts 的 key；解析失败时退化为 pdf.js 内部名（g_d0_f1） */
  font: string;
  text: string;
}

/**
 * 一张画在页面上的图片 —— PDF 里一次 image XObject 绘制，坐标是**外接矩形**。
 *
 * 与文字片段分开收：图片不经过 `getTextContent()`（那一路只吐 show-text 的产物），
 * 只能从算子表里连着 CTM 一起读出来。旋转过的图取外接矩形，正是 `wp:extent` 的语义
 * ——「用户拖出来的那个框」，所以两边可以直接比。
 */
export interface TruthImage {
  /** 外接矩形左上角 x（pt，页面左上角原点） */
  x: number;
  /** 外接矩形**顶边** y（pt，从页顶向下） */
  y: number;
  w: number;
  h: number;
  /**
   * 底边 y（= y + h）。冗余，但「图的底边坐在基线上没有」是图片这一类真值的头号问题，
   * 让它与 `TruthLine.y`（基线）能直接相减，比每次自己加一遍少一处出错的机会。
   */
  yBottom: number;
  /** pdf.js 的 XObject 名（`img_p0_1`）；同一张图重复引用时相同，用来认「这是第几张」 */
  name: string;
}

/** 按基线聚合出的一行 —— L1/L2 级断言（每页首末行、每行断行点）直接用这个 */
export interface TruthLine {
  /** 该行基线 y（pt） */
  y: number;
  /** 行左端 x、行右端 x（pt） */
  x: number;
  xEnd: number;
  text: string;
  /** 行首 / 行末字符（码点，非 UTF-16 单元），避头尾断言用 */
  first: string;
  last: string;
  /** 组成该行的片段在 page.items 中的下标 */
  items: number[];
}

export interface TruthPage {
  /** 0 基页序 */
  index: number;
  width: number;
  height: number;
  /** PDF 页面旋转角，正常公文恒为 0 */
  rotate: number;
  items: TruthItem[];
  lines: TruthLine[];
  /**
   * 本页画到的图片，按绘制顺序。**没有图片时整个字段不出现** —— 这样已入库的十几份
   * 纯文字真值重抽一遍仍然逐字节相同，加这个字段不会污染它们的 diff。
   */
  images?: TruthImage[];
}

export interface WordTruth {
  /** 源 docx 文件名 */
  source: string;
  generator: TruthGenerator;
  unit: 'pt';
  origin: 'top-left';
  pageCount: number;
  /** 全文用到的字体及其归一化度量 */
  fonts: Record<string, TruthFont>;
  /** Word COM 报告的页数；与 pageCount 不一致说明导出链路有问题 */
  wordPageCount?: number;
  sections?: WordSection[];
  pages: TruthPage[];
}
