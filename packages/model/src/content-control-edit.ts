/**
 * 普通文字命令也会改模板控件：提交前按最终显示同步元数据，避免 Word 重开时使用旧选择 / 日期。
 * 只看文字差异，格式调整、拆 run 与批注不会误清日期；不另加变更记录，连续输入仍能合并撤销。
 */

import type { ContentControl } from './content-control.ts';
import { contentControlSpans, textOfContentControl } from './content-control.ts';
import type { Body, NodeId } from './nodes.ts';
import { walkBlocks } from './nodes.ts';

/**
 * 外壳是否存在看成员标记，不能只看元数据表（被删控件的属性仍留在那里供回写与撤销使用）。
 * 块级外壳必须仍有块成员：合段后只剩行内后代引用它，不能冒充原来的块级容器。
 */
export function contentControlShells(body: Body, lockedOnly = false): Set<NodeId> {
  const controls = body.contentControls ?? {};
  const shellLocked = (c: ContentControl) => c.lock === 'sdtLocked' || c.lock === 'sdtContentLocked';
  const present = new Set<NodeId>();
  if (!Object.values(controls).some((c) => !lockedOnly || shellLocked(c))) return present;
  const collect = (inner: NodeId | undefined, scope: ContentControl['scope']) => {
    for (let id = inner, guard = 0; id !== undefined && guard < 64; guard++) {
      const control = controls[id];
      if (control === undefined) break;
      if (control.scope === scope && (!lockedOnly || shellLocked(control))) present.add(id);
      id = control.parent;
    }
  };
  for (const section of body.sections)
    for (const block of walkBlocks(section.blocks)) {
      collect(block.contentControl, 'block');
      if (block.kind === 'paragraph') for (const run of block.runs) collect(run.contentControl, 'inline');
    }
  return present;
}

export function reconcileContentControlText(
  before: Body,
  after: Body,
  filledText: ReadonlyMap<NodeId, string>,
): Body {
  const controls = after.contentControls;
  if (
    controls === undefined ||
    !Object.values(controls).some(
      (c) => c.type === 'date' || c.type === 'dropDownList' || c.type === 'comboBox',
    )
  )
    return after;
  const previous = new Map(contentControlSpans(before).map((s) => [s.control.id, textOfContentControl(s)]));
  let updated: Record<NodeId, ContentControl> | undefined;
  for (const span of contentControlSpans(after)) {
    const control = span.control;
    const choice = control.type === 'dropDownList' || control.type === 'comboBox';
    if (!choice && control.type !== 'date') continue;
    const text = textOfContentControl(span);
    // 显式填写刚确定的保存值 / 日期优先；后续若又打字，仍按打字后的最终内容同步。
    if (text === (filledText.get(control.id) ?? previous.get(control.id))) continue;
    const { dataBound: _dataBound, ...unbound } = control;
    let next = { ...unbound, showingPlaceholder: false };
    if (choice) {
      // 自由输入的文字是显示值，不能把恰好等于选项 value 的文字偷偷改成另一个选项。
      const item = control.items?.find((i) => i.text === text);
      const lastValue = item?.value ?? (control.type === 'comboBox' ? text : undefined);
      const { lastValue: _lastValue, ...rest } = next;
      next = lastValue === undefined ? rest : { ...rest, lastValue };
    } else if (control.date !== undefined) {
      // 任意文字无法可靠反推出日历日；宁可清属性，也不能让日期选择器指着旧日期。
      const { fullDate: _fullDate, ...date } = control.date;
      next = { ...next, date };
    }
    updated ??= { ...controls };
    updated[control.id] = next;
  }
  return updated === undefined ? after : { ...after, contentControls: updated };
}
