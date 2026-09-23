/**
 * 套用内建样式时补定义：Word 的「标题 1」在文档里用到之前，styles.xml 里往往根本没有它
 * （只在 `w:latentStyles` 里挂个名字）—— 点一下样式库，Word 才把整份定义写进去。
 * 我们照做：文档里有就用它（按 `w:name` 找，不按 id —— 中文版 Word 写的 id 是 `1` / `2`），
 * 没有才从这里补一份。
 *
 * 定义照**中文版 Word 默认模板**抄（Word 2016 起的 Normal.dotm）：
 * - 标题 1：二号（22pt）加粗、`w:kern` 22pt，段前 17pt、段后 16.5pt、2.41 倍行距（`w:line="578"`）
 * - 标题 2 / 3：三号（16pt）加粗，段前段后 13pt、1.73 倍行距（`w:line="416"`）；
 *   标题 2 用主题的标题字体（majorHAnsi / majorEastAsia），标题 3 不改字体
 * - 三级都 keepNext + keepLines，大纲级别 0 / 1 / 2，下一段样式回到正文
 * **这些数没有与真值对过**（没有「套用样式后导出」的 Word 样本），数值出自模板本身，
 * 不影响已有文档 —— 只有文档里本来没有这份样式、由我们补进去时才用得上。
 *
 * 纯函数、不改入参：定义挂在撤销快照上（`Body.styles`）。
 */
import type { ParaProps, RunProps } from './props.ts';
import type { StyleDefinition } from './styles.ts';

export type BuiltinStyleName = 'heading 1' | 'heading 2' | 'heading 3';

export const BUILTIN_STYLE_NAMES: readonly BuiltinStyleName[] = ['heading 1', 'heading 2', 'heading 3'];

const HEADING_FONTS: RunProps['fontThemes'] = {
  ascii: 'majorHAnsi',
  hAnsi: 'majorHAnsi',
  eastAsia: 'majorEastAsia',
  cs: 'majorBidi',
};

const PRESETS: Record<BuiltinStyleName, { id: string; paraProps: ParaProps; runProps: RunProps }> = {
  'heading 1': {
    id: 'Heading1',
    paraProps: {
      keepNext: true,
      keepLines: true,
      spacing: { before: 340, after: 330, line: 578, lineRule: 'auto' },
      outlineLevel: 0,
    },
    runProps: { bold: true, boldCs: true, kerning: 440, size: 440, sizeCs: 440 },
  },
  'heading 2': {
    id: 'Heading2',
    paraProps: {
      keepNext: true,
      keepLines: true,
      spacing: { before: 260, after: 260, line: 416, lineRule: 'auto' },
      outlineLevel: 1,
    },
    runProps: { fontThemes: HEADING_FONTS, bold: true, boldCs: true, size: 320, sizeCs: 320 },
  },
  'heading 3': {
    id: 'Heading3',
    paraProps: {
      keepNext: true,
      keepLines: true,
      spacing: { before: 260, after: 260, line: 416, lineRule: 'auto' },
      outlineLevel: 2,
    },
    runProps: { bold: true, boldCs: true, size: 320, sizeCs: 320 },
  },
};

/**
 * 一份内建样式的定义。`taken` 答某个 id 是否已被占用（文档里的样式与已补的定义都算）——
 * 英文版 Word 的 id 就是 `Heading1`，文档里可能有一份 id 撞上、名字却不是 heading 1 的
 * 自定义样式，撞了就加序号。`normalId` 是默认段落样式：basedOn 与下一段样式都指它。
 */
export function builtinStyleDefinition(
  name: BuiltinStyleName,
  taken: (id: string) => boolean,
  normalId: string,
): StyleDefinition {
  const preset = PRESETS[name];
  let id = preset.id;
  for (let n = 1; taken(id); n++) id = `${preset.id}_${n}`;
  return {
    id,
    name,
    ...(normalId === '' ? {} : { basedOn: normalId, next: normalId }),
    paraProps: structuredClone(preset.paraProps),
    runProps: structuredClone(preset.runProps),
    uiPriority: 9,
    quickFormat: true,
  };
}
