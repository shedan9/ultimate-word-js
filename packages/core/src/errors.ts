/**
 * 结构性错误。
 *
 * 与 `Diagnostic` 的分界是硬的（架构 §10）：**无法产出任何有意义结果**的才抛这里的错
 * —— 不是 zip、缺 document.xml、关系断裂。凡是「文档还能看，只是某处不对」的
 * （不认识的元素、basedOn 成环、字体缺失）一律记 Diagnostic 继续跑。
 *
 * 划错这条线的后果很具体：一份公文里有一个我们没实现的元素，用户应该看到文档，
 * 而不是白屏。
 */

/**
 * 错误码。用 const 对象而不是 enum —— `erasableSyntaxOnly` 禁止 enum，
 * 且字符串码在日志与上报里比数字可读。
 */
export const UwErrorCode = {
  /** 字节流根本不是 zip，或 zip 目录损坏 */
  NOT_A_ZIP: 'NOT_A_ZIP',
  /** 是 zip，但缺 OPC 必需部件（[Content_Types].xml / 根关系） */
  NOT_AN_OPC_PACKAGE: 'NOT_AN_OPC_PACKAGE',
  /** 是 OPC 包，但不是 WordprocessingML 文档（找不到 officeDocument 主部件） */
  NOT_A_WORD_DOCUMENT: 'NOT_A_WORD_DOCUMENT',
  /** 部件缺失：关系指向了一个包里不存在的 part */
  PART_NOT_FOUND: 'PART_NOT_FOUND',
  /** XML 无法解析 */
  MALFORMED_XML: 'MALFORMED_XML',
} as const;

export type UwErrorCode = (typeof UwErrorCode)[keyof typeof UwErrorCode];

export class UwError extends Error {
  readonly code: UwErrorCode;
  /** 出问题的部件名（如 `/word/document.xml`），没有具体部件时省略 */
  readonly part: string | undefined;

  constructor(code: UwErrorCode, message: string, options: { part?: string; cause?: unknown } = {}) {
    super(message, options.cause === undefined ? undefined : { cause: options.cause });
    this.name = 'UwError';
    this.code = code;
    this.part = options.part;
  }
}
