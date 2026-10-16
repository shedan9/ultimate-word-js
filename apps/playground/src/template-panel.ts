/** 宿主填写样例：按控件 id 提交，每次保存只改这一处，避免同名字段被一起覆盖。 */
import type { Disposable, DocRange, UwDocument, UwView } from 'ultimate-word';
import './template-panel.css';

const TYPES: Record<string, string> = {
  dropDownList: '下拉框',
  comboBox: '组合框',
  date: '日期',
  checkbox: '复选框',
};

export function createTemplatePanel(button: HTMLButtonElement, host: HTMLElement) {
  const dialog = document.createElement('dialog');
  dialog.className = 'template-panel';
  dialog.setAttribute('aria-labelledby', 'template-title');
  dialog.innerHTML = `
    <form novalidate>
      <div class="template-heading"><h2 id="template-title">填写模板</h2><button type="button" class="template-close">关闭</button></div>
      <p class="template-intro">选择字段，保存后更新文档中的这一处。</p>
      <label class="template-label" for="template-target">模板字段</label>
      <select id="template-target"></select>
      <div class="template-current"><span>当前值</span><p></p></div>
      <div class="template-fields"></div>
      <p class="template-message" id="template-message" role="status" aria-live="polite" aria-atomic="true"></p>
      <div class="template-actions">
        <div><button type="button" class="template-undo">撤销</button><button type="button" class="template-redo">重做</button></div>
        <div><button type="button" class="template-cancel">取消</button><button type="submit" class="template-save">保存到文档</button></div>
      </div>
    </form>`;
  host.append(dialog);
  const form = dialog.querySelector('form') as HTMLFormElement;
  const picker = dialog.querySelector('#template-target') as HTMLSelectElement;
  const fields = dialog.querySelector('.template-fields') as HTMLElement;
  const current = dialog.querySelector('.template-current p') as HTMLElement;
  const message = dialog.querySelector('.template-message') as HTMLElement;
  const save = dialog.querySelector('.template-save') as HTMLButtonElement;
  const undo = dialog.querySelector('.template-undo') as HTMLButtonElement;
  const redo = dialog.querySelector('.template-redo') as HTMLButtonElement;
  let doc: UwDocument | undefined;
  let view: UwView | undefined;
  let subscription: Disposable | undefined;
  let targetId = '';
  let applying = false;
  let choice: HTMLSelectElement | undefined;
  let input: HTMLInputElement | undefined;

  const notify = (text: string, error = false) => {
    message.textContent = text;
    message.classList.toggle('template-error', error);
    input?.setAttribute('aria-invalid', String(error));
    choice?.setAttribute('aria-invalid', String(error));
  };
  const target = () => doc?.bindings.list().find((b) => b.id === targetId);
  const label = (text: string, id: string) => {
    const el = document.createElement('label');
    el.className = 'template-label';
    el.htmlFor = id;
    el.textContent = text;
    fields.append(el);
  };
  const makeInput = (type: string) => {
    const el = document.createElement('input');
    el.type = type;
    el.id = 'template-value';
    el.setAttribute('aria-describedby', 'template-message');
    return el;
  };
  const renderValue = () => {
    const restoreFocus = fields.contains(document.activeElement);
    fields.replaceChildren();
    choice = undefined;
    input = undefined;
    const b = target();
    current.textContent = b?.placeholder ? `占位文字：${b.value}` : b?.value || '（空）';
    undo.disabled = !doc?.canUndo;
    redo.disabled = !doc?.canRedo;
    save.disabled = b === undefined || b.locked || TYPES[b.type] === undefined;
    notify('');
    if (b === undefined) {
      notify('此字段已被移除，请选择其他字段。');
      return;
    }
    if (b.type === 'dropDownList' || b.type === 'comboBox') {
      label('选择选项', 'template-choice');
      choice = document.createElement('select');
      choice.id = 'template-choice';
      choice.setAttribute('aria-describedby', 'template-message');
      choice.add(new Option('请选择选项', ''));
      (choice.options[0] as HTMLOptionElement).disabled = true;
      for (const [index, item] of (b.items ?? []).entries()) {
        const duplicate = (b.items ?? []).filter((other) => other.text === item.text).length > 1;
        choice.add(new Option(duplicate ? `${item.text}（${item.value}）` : item.text, String(index)));
      }
      if (b.type === 'comboBox') choice.add(new Option('自填文字', 'custom'));
      const match = b.placeholder
        ? -1
        : (b.items ?? []).findIndex(
            (i) => i.text === b.value && (b.lastValue === undefined || i.value === b.lastValue),
          );
      choice.value = match >= 0 ? String(match) : b.type === 'comboBox' && !b.placeholder ? 'custom' : '';
      choice.disabled = b.locked;
      fields.append(choice);
      if (b.type === 'comboBox') {
        const custom = document.createElement('div');
        custom.className = 'template-custom';
        input = makeInput('text');
        input.value = b.placeholder ? '' : b.value;
        input.disabled = b.locked;
        const caption = document.createElement('label');
        caption.className = 'template-label';
        caption.htmlFor = input.id;
        caption.textContent = '自填内容';
        custom.append(caption, input);
        custom.hidden = choice.value !== 'custom';
        choice.addEventListener('change', () => {
          custom.hidden = choice?.value !== 'custom';
          if (!custom.hidden) input?.focus();
          notify('');
        });
        fields.append(custom);
      }
      if (b.type === 'dropDownList' && !b.items?.length) {
        save.disabled = true;
        notify('此下拉框没有可选项，请先在原模板中配置选项。');
      }
    } else if (b.type === 'date') {
      label('选择日期', 'template-value');
      input = makeInput('date');
      input.min = '0001-01-01';
      input.max = '9999-12-31';
      input.value = b.fullDate?.slice(0, 10) ?? '';
      input.disabled = b.locked;
      fields.append(input);
    } else if (b.type === 'checkbox') {
      input = makeInput('checkbox');
      input.checked = b.checked ?? false;
      input.disabled = b.locked;
      const caption = document.createElement('label');
      caption.className = 'template-check';
      const text = document.createElement('span');
      text.textContent = '勾选此项';
      caption.append(input, text);
      fields.append(caption);
    } else {
      notify('此类型暂不支持面板填写。文字字段可切换到编辑模式填写。');
    }
    if (b.locked) notify('此字段或外层控件已锁定内容，无法填写。');
    if (restoreFocus) (choice ?? input)?.focus();
  };
  const refresh = () => {
    const list = doc?.bindings.list() ?? [];
    const counts = new Map<string, number>();
    const totals = new Map<string, number>();
    for (const b of list) totals.set(b.key, (totals.get(b.key) ?? 0) + 1);
    picker.replaceChildren();
    for (const b of list) {
      const index = (counts.get(b.key) ?? 0) + 1;
      counts.set(b.key, index);
      const suffix = (totals.get(b.key) ?? 0) > 1 ? `（第 ${index} 处）` : '';
      picker.add(new Option(`${b.alias || b.key}${suffix}${b.locked ? ' · 已锁定' : ''}`, b.id));
    }
    picker.value = targetId;
    renderValue();
  };
  const close = () => dialog.close();
  dialog.querySelector('.template-close')?.addEventListener('click', close);
  dialog.querySelector('.template-cancel')?.addEventListener('click', close);
  picker.addEventListener('change', () => {
    targetId = picker.value;
    renderValue();
    const range = target()?.range;
    if (range !== undefined) view?.scrollTo(range, { align: 'center' });
  });
  fields.addEventListener('input', () => {
    if (message.classList.contains('template-error')) notify('');
  });
  undo.addEventListener('click', () => doc?.undo());
  redo.addEventListener('click', () => doc?.redo());
  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const b = target();
    if (doc === undefined || b === undefined || save.disabled) return;
    try {
      applying = true;
      let filled: DocRange | undefined;
      doc.tx((t) => {
        if (b.type === 'checkbox') filled = t.setContentControlChecked(b.id, input?.checked ?? false);
        else if (b.type === 'date') {
          if (!input?.value || !input.checkValidity()) throw new Error('请选择有效日期。');
          filled = t.setContentControlDate(b.id, input.value);
        } else if (choice?.value === 'custom') filled = t.fillContentControl(b.id, input?.value ?? '');
        else {
          if (!choice?.value) throw new Error('请选择一个选项。');
          filled = t.selectContentControlItem(b.id, Number(choice.value));
        }
      });
      close();
      if (filled !== undefined) view?.scrollTo(filled, { align: 'center' });
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      notify(
        reason.includes('dateFormat') ? '此日期字段缺少显示格式，请先在原模板中设置格式。' : reason,
        true,
      );
    } finally {
      applying = false;
    }
  });
  const open = (id?: string): boolean => {
    const list = doc?.bindings.list() ?? [];
    const b =
      id === undefined
        ? (list.find((b) => TYPES[b.type] !== undefined && !b.locked) ?? list[0])
        : list.find((b) => b.id === id);
    if (b === undefined) return false;
    targetId = b.id;
    refresh();
    if (!dialog.open) dialog.showModal();
    picker.focus();
    return true;
  };
  button.addEventListener('click', () => open());
  return {
    open,
    attach(nextDoc: UwDocument, nextView: UwView) {
      if (dialog.open) close();
      subscription?.dispose();
      doc = nextDoc;
      view = nextView;
      button.disabled = doc.bindings.list().length === 0;
      subscription = doc.on('document:change', () => {
        button.disabled = doc?.bindings.list().length === 0;
        if (dialog.open && !applying) {
          refresh();
          if (!message.textContent) notify('文档已更新，已载入最新值。');
        }
      });
    },
  };
}
