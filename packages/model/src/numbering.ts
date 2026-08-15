/**
 * `numbering.xml` → 编号定义。
 *
 * Phase 1 只**解析**，不消费：编号那一层要插进级联（在段落样式之后、直接格式之前），
 * 还要按文档顺序累计计数器、按 `w:lvlText` 生成编号文字 —— 那是 Phase 5 的活。
 * 现在先把数据结构立起来，因为它是级联里那个写明的洞的另一半。
 *
 * 两层间接不能省：
 *
 * ```
 * 段落 w:numPr{numId, ilvl} → w:num[numId] → w:abstractNum[abstractNumId] → w:lvl[ilvl]
 *                                    └── w:lvlOverride 可以就地改某一级
 * ```
 *
 * 直接拿 `numId` 当编号定义用是最常见的错法：同一个 abstractNum 会被多个 num 引用，
 * 「重新开始编号」正是靠不同的 num + `w:startOverride` 实现的。
 */
import type { DiagnosticSink } from '@uw/core';
import type { XmlDocument, XmlElement } from '@uw/ooxml';
import { child, children } from '@uw/ooxml';
import { parseParaProps, parseRunProps } from './parse-props.ts';
import type { Justification, ParaProps, RunProps } from './props.ts';
import { attrInt, attrOf, enumVal, intVal } from './xml-values.ts';

const SUFFIXES = ['tab', 'space', 'nothing'] as const;
const JUSTIFICATIONS = ['left', 'center', 'right'] as const;

export interface NumberingLevel {
  /** `w:ilvl`，0 起 */
  level: number;
  /** `w:start`：本级从几开始 */
  start: number;
  /**
   * `w:numFmt`：编号形式。**故意留成字符串**，不收成联合类型 ——
   * 规范里有六十多种，中文文档常用的 `chineseCounting`（一二三）、
   * `chineseLegalSimplified`（壹贰叁）、`ideographDigital`（１２３）都在其中，
   * 收窄只会在遇到没列的值时把编号整个丢掉。认不出的值留给 Phase 5 降级成 decimal。
   */
  numFmt: string;
  /** `w:lvlText`：`%1.` 这种模板，`%n` 引用第 n 级的计数值（**1 起，不是 0 起**） */
  lvlText: string;
  justification: Justification;
  /** 编号与正文之间用什么分隔。缺省是制表位，不是空格 */
  suffix: 'tab' | 'space' | 'nothing';
  /** `w:lvlRestart`：本级在第几级变化时归零。0 = 从不归零；缺省 = 上一级变就归零 */
  restartAfter?: number;
  /** `w:isLgl`：所有层级一律用阿拉伯数字显示（法律文书编号） */
  isLegal: boolean;
  /** `w:pStyle`：本级绑定的段落样式 */
  paraStyleId?: string;
  /** 本级自带的段落属性（几乎总有缩进）与字符属性（项目符号的字体） */
  paraProps: ParaProps;
  runProps: RunProps;
}

export interface AbstractNumbering {
  id: number;
  /** `w:multiLevelType`：singleLevel / multilevel / hybridMultilevel */
  multiLevelType: string;
  levels: Record<number, NumberingLevel>;
  /**
   * `w:numStyleLink`：这个 abstractNum 自己没有级别定义，真正的定义在某个编号样式上。
   * 见文件末尾 —— Phase 5 要跟着跳一次。
   */
  numStyleLink?: string;
  /** `w:styleLink`：反过来，声明「本定义被某个编号样式使用」 */
  styleLink?: string;
}

export interface LevelOverride {
  /** `w:startOverride`：本实例这一级从几开始（「重新编号」就是它） */
  start?: number;
  /** `w:lvlOverride` 里可以整个替换一级的定义 */
  level?: NumberingLevel;
}

export interface NumberingInstance {
  numId: number;
  abstractNumId: number;
  overrides: Record<number, LevelOverride>;
}

export interface Numbering {
  abstract: Record<number, AbstractNumbering>;
  instances: Record<number, NumberingInstance>;
}

export const EMPTY_NUMBERING: Numbering = { abstract: {}, instances: {} };

export function parseNumbering(doc: XmlDocument | undefined, diagnostics: DiagnosticSink): Numbering {
  if (doc === undefined) return structuredClone(EMPTY_NUMBERING);
  const out: Numbering = { abstract: {}, instances: {} };

  for (const el of children(doc.root, 'w:abstractNum')) {
    const id = attrInt(el, 'w:abstractNumId');
    if (id === undefined) continue;
    out.abstract[id] = parseAbstract(el, id);
  }

  for (const el of children(doc.root, 'w:num')) {
    const numId = attrInt(el, 'w:numId');
    const abstractNumId = intVal(el, 'w:abstractNumId');
    if (numId === undefined || abstractNumId === undefined) continue;
    if (out.abstract[abstractNumId] === undefined) {
      diagnostics.warn(
        'numbering-missing-abstract',
        `w:num ${numId} 指向不存在的 abstractNum ${abstractNumId}，该段落将没有编号`,
        { part: 'numbering.xml' },
      );
    }
    out.instances[numId] = { numId, abstractNumId, overrides: parseOverrides(el) };
  }
  return out;
}

function parseAbstract(el: XmlElement, id: number): AbstractNumbering {
  const levels: Record<number, NumberingLevel> = {};
  for (const lvl of children(el, 'w:lvl')) {
    const parsed = parseLevel(lvl);
    if (parsed !== undefined) levels[parsed.level] = parsed;
  }
  const out: AbstractNumbering = {
    id,
    multiLevelType: attrOf(child(el, 'w:multiLevelType'), 'w:val') ?? 'multilevel',
    levels,
  };
  const numStyleLink = attrOf(child(el, 'w:numStyleLink'), 'w:val');
  const styleLink = attrOf(child(el, 'w:styleLink'), 'w:val');
  if (numStyleLink !== undefined) out.numStyleLink = numStyleLink;
  if (styleLink !== undefined) out.styleLink = styleLink;
  return out;
}

function parseLevel(lvl: XmlElement): NumberingLevel | undefined {
  const level = attrInt(lvl, 'w:ilvl');
  if (level === undefined) return undefined;

  const out: NumberingLevel = {
    level,
    // 没写 w:start 时从 1 开始，不是 0
    start: intVal(lvl, 'w:start') ?? 1,
    numFmt: attrOf(child(lvl, 'w:numFmt'), 'w:val') ?? 'decimal',
    lvlText: attrOf(child(lvl, 'w:lvlText'), 'w:val') ?? '',
    justification: enumVal(attrOf(child(lvl, 'w:lvlJc'), 'w:val'), JUSTIFICATIONS) ?? 'left',
    suffix: enumVal(attrOf(child(lvl, 'w:suff'), 'w:val'), SUFFIXES) ?? 'tab',
    isLegal: child(lvl, 'w:isLgl') !== undefined,
    paraProps: parseParaProps(child(lvl, 'w:pPr')),
    runProps: parseRunProps(child(lvl, 'w:rPr')),
  };
  const restartAfter = intVal(lvl, 'w:lvlRestart');
  const paraStyleId = attrOf(child(lvl, 'w:pStyle'), 'w:val');
  if (restartAfter !== undefined) out.restartAfter = restartAfter;
  if (paraStyleId !== undefined) out.paraStyleId = paraStyleId;
  return out;
}

function parseOverrides(num: XmlElement): Record<number, LevelOverride> {
  const out: Record<number, LevelOverride> = {};
  for (const el of children(num, 'w:lvlOverride')) {
    const ilvl = attrInt(el, 'w:ilvl');
    if (ilvl === undefined) continue;
    const o: LevelOverride = {};
    const start = intVal(el, 'w:startOverride');
    if (start !== undefined) o.start = start;
    const lvl = child(el, 'w:lvl');
    const parsed = lvl === undefined ? undefined : parseLevel(lvl);
    if (parsed !== undefined) o.level = parsed;
    out[ilvl] = o;
  }
  return out;
}

/**
 * `numId` + `ilvl` → 最终的级别定义，`w:lvlOverride` 已经叠好。
 *
 * **`numId = 0` 是「取消编号」**，不是「第 0 号编号」：段落样式给了编号、
 * 而这一段想去掉时，Word 写的就是 `<w:numId w:val="0"/>`。当成正常编号去查会
 * 给本该没编号的段落加上编号。
 */
export function numberingLevel(n: Numbering, numId: number, ilvl: number): NumberingLevel | undefined {
  if (numId === 0) return undefined;
  const instance = n.instances[numId];
  if (instance === undefined) return undefined;

  const override = instance.overrides[ilvl];
  const base = override?.level ?? n.abstract[instance.abstractNumId]?.levels[ilvl];
  if (base === undefined) return undefined;
  // startOverride 只改起始值，其余照旧
  return override?.start === undefined ? base : { ...base, start: override.start };
}

// ── 写明的洞：w:numStyleLink ──────────────────────────────────────────────────
//
// 一个 abstractNum 可以只写 `<w:numStyleLink w:val="某编号样式"/>` 而没有任何 w:lvl，
// 真正的级别定义在那个样式的 `w:pPr/w:numPr/w:numId` 指向的**另一个** num 上。
// 也就是说完整的解引用要多跳一次，而且可能成环。
//
// 这里**没有**跟这一跳：Phase 1 不消费编号，跳了也没人用；而要跳就必须一并处理环检测，
// 那些代码现在无法被任何测试覆盖。Phase 5 接编号时，入口就是 `numberingLevel`：
// base 为 undefined 且 abstract 有 numStyleLink 时，去 styles 里找同名样式再跳一次，
// 带上和 `styles.ts/chainOf` 一样的 seen 集合防环。
