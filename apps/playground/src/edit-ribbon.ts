import { ptToTwips, twipsToPt } from '@uw/core';
import type { FormattingCommand, FormattingState, UwDocument, UwView } from 'ultimate-word';
import './edit-ribbon.css';

const paths = {
  undo: 'M8 4 3 9l5 5 M4 9h9a5 5 0 0 1 0 10',
  redo: 'm16 4 5 5-5 5 M20 9h-9a5 5 0 0 0 0 10',
  bullet: 'M3 5h1m-1 6h1m-1 6h1 M8 5h13M8 11h13M8 17h13',
  decimal: 'M2 4l2-1v5M2 8h4 M2 14q4-3 4 0l-4 4h4 M9 5h12M9 11h12M9 17h12',
  in: 'M3 4h18M11 9h10M11 14h10M3 19h18 M3 8l4 4-4 4',
  out: 'M3 4h18M11 9h10M11 14h10M3 19h18 M7 8l-4 4 4 4',
  left: 'M3 4h18M3 9h11M3 14h18M3 19h11',
  center: 'M3 4h18M7 9h10M3 14h18M7 19h10',
  right: 'M3 4h18M10 9h11M3 14h18M10 19h11',
  both: 'M3 4h18M3 9h18M3 14h18M3 19h18',
  clear: 'm4 13 8-9 7 6-8 9H9z M8 9l7 6 M10 19h11',
};
function icon(name: keyof typeof paths): string {
  return `<svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round" stroke-linecap="round"><path d="${paths[name]}"/></svg>`;
}

export function createEditRibbon(host: HTMLElement) {
  const ribbon = document.createElement('section');
  ribbon.id = 'edit-ribbon';
  ribbon.hidden = true;
  ribbon.setAttribute('aria-label', '开始编辑工具栏');
  ribbon.innerHTML = `
    <div class="ribbon-heading"><strong>开始</strong><span class="ribbon-history" role="group" aria-label="编辑历史"></span></div>
    <div class="ribbon-groups">
      <div class="ribbon-group" role="group" aria-label="字体">
        <div class="ribbon-row"><select aria-label="字体"></select><select aria-label="字号"></select></div>
        <div class="ribbon-row ribbon-character"></div><span class="ribbon-caption">字体</span>
      </div>
      <div class="ribbon-group" role="group" aria-label="段落">
        <div class="ribbon-row ribbon-list"></div><div class="ribbon-row ribbon-align"></div><span class="ribbon-caption">段落</span>
      </div>
      <div class="ribbon-group" role="group" aria-label="样式">
        <div class="ribbon-gallery"></div><span class="ribbon-caption">样式</span>
      </div>
    </div>`;
  host.append(ribbon);
  const slot = (selector: string) => ribbon.querySelector<HTMLElement>(selector) as HTMLElement;
  let view: UwView | undefined;
  let doc: UwDocument | undefined;
  const reflect: ((state: FormattingState | undefined) => void)[] = [];
  function button(parent: HTMLElement, label: string, content: string, click: () => void) {
    const el = document.createElement('button');
    el.type = 'button';
    el.title = label;
    el.setAttribute('aria-label', label);
    el.innerHTML = content;
    el.addEventListener('click', () => {
      click();
      view?.focus?.();
    });
    parent.append(el);
    return el;
  }
  function command(
    parent: HTMLElement,
    label: string,
    content: string,
    cmd: FormattingCommand,
    pressed?: (state: FormattingState) => boolean | undefined,
  ) {
    const el = button(parent, label, content, () => view?.formatting?.execute(cmd));
    reflect.push((state) => {
      el.disabled = !state;
      if (pressed) el.setAttribute('aria-pressed', state ? String(pressed(state) ?? 'mixed') : 'false');
    });
  }
  // 鼠标点按钮时保留正文焦点与选区；下拉框与键盘仍使用原生焦点行为。
  ribbon.addEventListener('mousedown', (event) => {
    if (event.target instanceof Element && event.target.closest('button')) event.preventDefault();
  });
  const history = slot('.ribbon-history');
  const undo = button(history, '撤销', icon('undo'), () => doc?.undo());
  const redo = button(history, '重做', icon('redo'), () => doc?.redo());

  const character = slot('.ribbon-character');
  for (const [value, label, glyph] of [
    ['bold', '加粗', '<b>B</b>'],
    ['italic', '倾斜', '<i>I</i>'],
    ['underline', '下划线', '<u>U</u>'],
    ['subscript', '下标', 'x<sub>2</sub>'],
    ['superscript', '上标', 'x<sup>2</sup>'],
  ] as const)
    command(character, label, glyph, { type: 'toggle', value }, (s) => s[value]);
  command(character, '清除字符格式', icon('clear'), { type: 'clearCharacter' });
  const color = document.createElement('label');
  color.className = 'ribbon-color';
  color.title = '字体颜色';
  color.innerHTML = '<b aria-hidden="true">A</b><input type="color" aria-label="字体颜色">';
  character.append(color);
  const picker = color.querySelector('input') as HTMLInputElement;
  picker.addEventListener('change', () => {
    view?.formatting?.execute({
      type: 'character',
      patch: { color: picker.value.slice(1).toUpperCase(), themeColor: null },
    });
    view?.focus?.();
  });
  const fonts = ribbon.querySelector('select[aria-label="字体"]') as HTMLSelectElement;
  const sizes = ribbon.querySelector('select[aria-label="字号"]') as HTMLSelectElement;
  for (const font of [
    '宋体',
    '仿宋',
    '楷体',
    '黑体',
    '微软雅黑',
    '等线',
    'Calibri',
    'Arial',
    'Times New Roman',
  ])
    fonts.add(new Option(font));
  for (const [size, label] of [
    [9, '小五'],
    [10.5, '五号'],
    [12, '小四'],
    [14, '四号'],
    [15, '小三'],
    [16, '三号'],
    [18, '小二'],
    [22, '二号'],
    [24, '小一'],
    [26, '一号'],
    [36, '小初'],
    [42, '初号'],
  ] as const) {
    sizes.add(new Option(`${label} · ${size}`, String(size)));
  }
  fonts.addEventListener('change', () => {
    const family = fonts.value;
    view?.formatting?.execute({
      type: 'character',
      patch: { fonts: { eastAsia: family, ascii: family, hAnsi: family, cs: family }, fontThemes: null },
    });
    view?.focus?.();
  });
  sizes.addEventListener('change', () => {
    const size = ptToTwips(Number(sizes.value));
    view?.formatting?.execute({ type: 'character', patch: { size, sizeCs: size } });
    view?.focus?.();
  });
  function selectValue(select: HTMLSelectElement, value: string) {
    // 文件可能使用清单外的字体 / 字号。用 Option 写纯文本，文档里的名称不进入 HTML。
    if (value && !Array.from(select.options).some((o) => o.value === value)) select.add(new Option(value));
    select.value = value;
  }
  const list = slot('.ribbon-list');
  for (const [value, label] of [
    ['bullet', '项目符号'],
    ['decimal', '编号'],
  ] as const)
    command(list, label, icon(value), { type: 'list', value }, (s) => s.list === value);
  for (const [value, label] of [
    ['out', '减少缩进'],
    ['in', '增加缩进'],
  ] as const)
    command(list, label, icon(value), { type: 'indent', value });
  const alignment = slot('.ribbon-align');
  for (const [value, label] of [
    ['left', '左对齐'],
    ['center', '居中'],
    ['right', '右对齐'],
    ['both', '两端对齐'],
  ] as const)
    command(alignment, label, icon(value), { type: 'align', value }, (s) => s.alignment === value);
  const gallery = slot('.ribbon-gallery');
  for (const [value, label] of [
    ['normal', '正文'],
    ['heading 1', '标题 1'],
    ['heading 2', '标题 2'],
    ['heading 3', '标题 3'],
  ] as const)
    command(
      gallery,
      label,
      `<span aria-hidden="true">Aa</span>${label}`,
      { type: 'style', value },
      (s) => s.style === value,
    );

  function refresh() {
    const state = view?.formatting?.read();
    for (const update of reflect) update(state);
    const unavailable = view?.selection !== undefined && !state;
    undo.disabled = unavailable || !doc?.canUndo;
    redo.disabled = unavailable || !doc?.canRedo;
    fonts.disabled = sizes.disabled = picker.disabled = !state;
    selectValue(fonts, state?.font ?? '');
    selectValue(sizes, state?.size ? String(twipsToPt(state.size)) : '');
    picker.value = state?.color && /^[\da-f]{6}$/i.test(state.color) ? `#${state.color}` : '#000000';
    color.style.setProperty('--ink-color', picker.value);
  }
  return {
    attach(opened: UwDocument, mounted: UwView) {
      doc = opened;
      view = mounted;
      ribbon.hidden = !view.formatting;
      if (ribbon.hidden) return;
      view.on('format:change', refresh);
      refresh();
    },
  };
}
