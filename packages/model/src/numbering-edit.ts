/**
 * 新建列表：往编号定义里加一份 `w:abstractNum` + 一个指向它的 `w:num`。
 *
 * 每次新建都是**新的一对**，不复用文档里已有的 abstractNum：计数按 numId 分家
 * （numbering-counter.ts 规则 1），而「样子相同」的判断要逐级比九层定义，比错了会把
 * 两个本该各数各的列表接成一个。要接着上一个列表数，调用方直接用它的 numId。
 *
 * 纯函数、不改入参：定义挂在撤销快照上（`Body.numbering`），被冻结的旧快照必须原样留着。
 */
import type { AbstractNumbering, Numbering, NumberingLevel } from './numbering.ts';

export type ListKind = 'bullet' | 'decimal';

/** Word 的列表只有 0–8 九级（`w:ilvl`）。 */
export const LIST_LEVELS = 9;

/**
 * 每级缩进 420 twips、悬挂 420 —— 中文版 Word「编号库 / 项目符号库」的默认值
 * （五号字两个字宽），逐级右移一格。
 *
 * 编号三级一轮：`1.` → `a)` → `i.`，与中文版 Word 的默认编号一致。
 * 项目符号**没有照抄** Word：它用 Wingdings 的私用区码点（U+F06C 等），
 * 我们既没有这款字体的度量也没有符号字体到 Unicode 的映射，照抄只会画出豆腐块。
 * 改用 Unicode 里同形的 ● ○ ■，回写后 Word 打开也显示得出来。
 * 编号自身对齐（`w:lvlJc`）一律 left：Word 的罗马数字级是 right，但布局还不认 lvlJc，
 * 写了也不生效，不摆不生效的值。
 */
const DECIMAL_CYCLE = [
  { numFmt: 'decimal', text: (l: number) => `%${l + 1}.` },
  { numFmt: 'lowerLetter', text: (l: number) => `%${l + 1})` },
  { numFmt: 'lowerRoman', text: (l: number) => `%${l + 1}.` },
] as const;
const BULLETS = ['●', '○', '■'] as const;
const INDENT_STEP = 420;

function presetLevel(kind: ListKind, level: number): NumberingLevel {
  const cycle = DECIMAL_CYCLE[level % DECIMAL_CYCLE.length] as (typeof DECIMAL_CYCLE)[number];
  return {
    level,
    start: 1,
    numFmt: kind === 'bullet' ? 'bullet' : cycle.numFmt,
    lvlText: kind === 'bullet' ? (BULLETS[level % BULLETS.length] as string) : cycle.text(level),
    justification: 'left',
    suffix: 'tab',
    isLegal: false,
    paraProps: { indent: { left: INDENT_STEP * (level + 1), hanging: INDENT_STEP } },
    runProps: {},
  };
}

/** 返回加了定义的新 `Numbering` 与新 num 的 id。id 取现有最大值 + 1，numId 从 1 起（0 是「取消编号」）。 */
export function addListDefinition(
  numbering: Numbering,
  kind: ListKind,
): { numbering: Numbering; numId: number } {
  const next = (ids: string[], min: number) => Math.max(min - 1, ...ids.map(Number)) + 1;
  const abstractId = next(Object.keys(numbering.abstract), 0);
  const numId = next(Object.keys(numbering.instances), 1);
  const levels: Record<number, NumberingLevel> = {};
  for (let l = 0; l < LIST_LEVELS; l++) levels[l] = presetLevel(kind, l);
  const abstract: AbstractNumbering = {
    id: abstractId,
    multiLevelType: 'hybridMultilevel',
    levels,
  };
  return {
    numbering: {
      abstract: { ...numbering.abstract, [abstractId]: abstract },
      instances: { ...numbering.instances, [numId]: { numId, abstractNumId: abstractId, overrides: {} } },
    },
    numId,
  };
}
