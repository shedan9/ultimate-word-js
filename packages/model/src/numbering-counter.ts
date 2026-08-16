/**
 * 编号计数器 —— 编号里唯一**有状态**的一环。
 *
 * 别处的一切（解析、格式化、级联）都是纯函数，只有这里必须按**文档顺序**一段一段喂：
 * 第 3 段的编号是几，取决于前面出现过哪些段。所以它是个显式的状态对象，
 * 由 `resolveBody()` 持有并沿着遍历顺序推进 —— 而不是藏在某个模块级变量里，
 * 那样第二次加载文档就会接着上一份的计数往下数。
 *
 * 三条规则决定了「第几」：
 *
 * 1. **计数按 `numId` 分家，不按 abstractNumId。** 两个 `w:num` 指向同一份
 *    `w:abstractNum` 是「同样的样子、各数各的」，公文里「一、二、三」重新起头就靠它。
 *    按 abstractNumId 计数会让第二个列表接着第一个数下去
 * 2. **某级递增时，更深的级归零。** 归零的范围由 `w:lvlRestart` 收窄：
 *    0 = 从不归零；n = 只有第 n 级（含更浅的）动了才归零
 * 3. **跳级引用取 start，且不推进。** 文档里第一段就是 ilvl=1 时显示「1.1」，
 *    但第 0 级并没有被用掉 —— 后面真出现 ilvl=0 的段落时它仍是 1
 */
import { formatLevelText } from './number-format.ts';
import type { Numbering, NumberingLevel, StyleLookup } from './numbering.ts';
import { numberingLevel } from './numbering.ts';

export interface NumberedParagraph {
  /** 该级定义（`w:lvlOverride` 已叠好），编号的 pPr / rPr 从这儿取 */
  level: NumberingLevel;
  /** 展开后的编号文字：`一、` / `1.1` / `` 。`numFmt=none` 时是空串 */
  text: string;
  /** 本段在本级的计数值。交叉引用（Phase 5 的域）要的是这个数，不是那串字 */
  value: number;
}

export interface NumberingCounters {
  /**
   * 一个段落用到了 `(numId, ilvl)`：推进计数器并返回该段的编号。
   *
   * **有副作用，每个段落只能调一次**，而且调用顺序必须是文档顺序。
   * 拿不到级别定义（numId=0 的「取消编号」、指向不存在的定义）时返回 undefined，
   * 且**不推进任何计数器** —— 一个没编号的段落不该把后面的编号顶掉一位。
   */
  advance(numId: number, ilvl: number): NumberedParagraph | undefined;
}

export function createNumberingCounters(n: Numbering, styles?: StyleLookup): NumberingCounters {
  /** numId → (ilvl → 当前值)。缺席 = 这一级还没出现过，下次取 start */
  const state = new Map<number, Map<number, number>>();
  const levelOf = (numId: number, ilvl: number): NumberingLevel | undefined =>
    numberingLevel(n, numId, ilvl, styles);

  return {
    advance(numId, ilvl) {
      const level = levelOf(numId, ilvl);
      if (level === undefined) return undefined;

      let counters = state.get(numId);
      if (counters === undefined) {
        counters = new Map();
        state.set(numId, counters);
      }

      const previous = counters.get(ilvl);
      const value = previous === undefined ? level.start : previous + 1;
      counters.set(ilvl, value);
      resetDeeper(counters, ilvl, numId, levelOf);

      const values: number[] = [];
      const formats: string[] = [];
      for (let i = 0; i <= ilvl; i++) {
        const def = i === ilvl ? level : levelOf(numId, i);
        // 跳级时上层没值：按 start 显示但**不写回** counters，见文件头规则 3
        values.push(counters.get(i) ?? def?.start ?? 1);
        // w:isLgl：本级显示时所有层级一律阿拉伯数字（法律文书的「第 1.2 条」）
        formats.push(level.isLegal ? 'decimal' : (def?.numFmt ?? 'decimal'));
      }

      return { level, text: labelText(level, values, formats), value };
    },
  };
}

/**
 * `w:numFmt` 为 bullet / none 时 `w:lvlText` 不是模板。
 *
 * 项目符号的 lvlText 是字面字符（实心圆点是 Symbol 字体的 U+F0B7），里面真要出现
 * `%1` 也是原样打出来的字面量。拿模板去展开会把它替换成一个数字。
 */
function labelText(level: NumberingLevel, values: number[], formats: string[]): string {
  if (level.numFmt === 'none') return '';
  if (level.numFmt === 'bullet') return level.lvlText;
  return formatLevelText(level.lvlText, values, formats);
}

/**
 * 更深的级归零。
 *
 * `w:lvlRestart` 统一成一个阈值：不写 = 本级自己的级号（任何更浅的级动了都归零，
 * 也就是默认行为），写 0 = 永不归零，写 n = 只有第 n 级（1 起）或更浅的级动了才归零。
 * 这样三种情况是同一个比较，不必分支。
 */
function resetDeeper(
  counters: Map<number, number>,
  changed: number,
  numId: number,
  levelOf: (numId: number, ilvl: number) => NumberingLevel | undefined,
): void {
  for (const ilvl of [...counters.keys()]) {
    if (ilvl <= changed) continue;
    const threshold = levelOf(numId, ilvl)?.restartAfter ?? ilvl;
    if (threshold > 0 && changed <= threshold - 1) counters.delete(ilvl);
  }
}
