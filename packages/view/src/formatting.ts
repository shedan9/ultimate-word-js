import type { BuiltinStyleName, Justification, ListKind, RunPropsPatch } from '@uw/model';
import type { EditingController, ToggleFormat } from './editing.ts';
import type { EditingBinding } from './editing-dom.ts';

export type FormattingCommand =
  | { type: 'toggle'; value: ToggleFormat }
  | { type: 'character'; patch: RunPropsPatch }
  | { type: 'clearCharacter' }
  | { type: 'align'; value: Justification }
  | { type: 'list'; value: ListKind }
  | { type: 'indent'; value: 'in' | 'out' }
  | { type: 'style'; value: BuiltinStyleName | 'normal' };

export interface ViewFormatting {
  execute(command: FormattingCommand): void;
  read(): FormattingState | undefined;
}

export interface FormattingState {
  font: string | undefined;
  /** 与模型相同，单位为 twips。 */
  size: number | undefined;
  color: string | undefined;
  bold: boolean | undefined;
  italic: boolean | undefined;
  underline: boolean | undefined;
  superscript: boolean | undefined;
  subscript: boolean | undefined;
  alignment: Justification | undefined;
  list: ListKind | 'none' | undefined;
  style: string | undefined;
}

/** 混合值用 undefined 表示，避免把选区第一个字的格式误报为整段一致。 */
function common<T, V>(values: readonly T[], read: (item: T) => V): V | undefined {
  const first = values[0];
  if (first === undefined) return undefined;
  const value = read(first);
  return values.every((item) => read(item) === value) ? value : undefined;
}

export function selectionFormatting(
  state: EditingController,
  binding: EditingBinding,
): FormattingState | undefined {
  const range = state.selection;
  if (!range || state.composing) return undefined;
  const runs = binding.format?.(range, state.pendingFormat) ?? [];
  const paragraphs = binding.paragraphFormat?.(range) ?? [];
  return {
    font: common(runs, (r) => r.fonts?.eastAsia || r.fonts?.ascii),
    color: common(runs, (r) => r.color),
    size: common(runs, (r) => r.size),
    bold: common(runs, (r) => r.bold),
    italic: common(runs, (r) => r.italic),
    underline: common(runs, (r) => !!r.underline && r.underline !== 'none'),
    superscript: common(runs, (r) => r.vertAlign === 'superscript'),
    subscript: common(runs, (r) => r.vertAlign === 'subscript'),
    alignment: common(paragraphs, (p) => p.props.justification),
    list: common(paragraphs, (p) => {
      const label = p.props.numbering.label;
      return label ? (label.format === 'bullet' ? 'bullet' : 'decimal') : 'none';
    }),
    style: common(paragraphs, (p) => {
      const style =
        binding.editor.body.styles?.find((s) => s.id === p.styleId) ??
        binding.styles?.find((s) => s.id === p.styleId);
      if (!p.styleId || (style && 'isDefault' in style && style.isDefault)) return 'normal';
      return style?.name.toLowerCase() ?? p.styleId;
    }),
  };
}

export function applyFormatting(state: EditingController, command: FormattingCommand): void {
  switch (command.type) {
    case 'toggle':
      state.toggleFormat(command.value);
      break;
    case 'character':
      state.setRunFormat(command.patch);
      break;
    case 'clearCharacter':
      state.clearRunFormat();
      break;
    case 'align':
      state.align(command.value);
      break;
    case 'list':
      state.toggleList(command.value);
      break;
    case 'indent':
      state.indent(command.value);
      break;
    case 'style':
      state.applyBuiltinStyle(command.value);
      break;
  }
}
