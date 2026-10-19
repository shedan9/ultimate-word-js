import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, rangeOfContentControl } from './content-control.ts';
import type { Body } from './nodes.ts';
import { parseBody } from './parse-body.ts';
import type { TextTransaction } from './text-transaction.ts';
import { createTextEditor } from './text-transaction.ts';

const cases = [
  {
    type: 'dropDownList',
    pr: '<w:dropDownList w:lastValue="2"><w:listItem w:displayText="一般" w:value="2"/></w:dropDownList>',
    content: '<w:r><w:rPr><w:b/></w:rPr><w:t>一</w:t></w:r><w:r><w:t>般</w:t></w:r>',
  },
  {
    type: 'comboBox',
    pr: '<w:comboBox w:lastValue="2"><w:listItem w:displayText="一般" w:value="2"/></w:comboBox>',
    content: '<w:r><w:rPr><w:b/></w:rPr><w:t>一</w:t></w:r><w:r><w:t>般</w:t></w:r>',
  },
  {
    type: 'date',
    pr: '<w:date w:fullDate="2026-10-13T00:00:00Z"><w:dateFormat w:val="yyyy年M月d日"/></w:date>',
    content: '<w:r><w:rPr><w:b/></w:rPr><w:t>2026年</w:t></w:r><w:r><w:t>10月13日</w:t></w:r>',
  },
  { type: 'checkbox', pr: '<w14:checkbox/>', content: '<w:r><w:t>☐</w:t></w:r>' },
  {
    type: 'checkboxSymbol',
    pr: '<w14:checkbox/>',
    content: '<w:r><w:rPr><w:b/></w:rPr><w:sym w:font="MS Gothic" w:char="2610"/></w:r>',
  },
] as const;

function parse(pr: string, content: string, scope: string): Body {
  const inner = scope === 'block' ? `<w:p>${content}</w:p>` : content;
  const control = `<w:sdt><w:sdtPr><w:tag w:val="field"/><w:dataBinding w:xpath="/field"/>${pr}</w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
  return parseBody(
    parseXml(
      `<w:document><w:body>${scope === 'block' ? control : `<w:p>${control}</w:p>`}</w:body></w:document>`,
    ),
    createDiagnosticSink(),
  );
}

function fill(t: TextTransaction, type: string, id: string) {
  if (type === 'date') return t.setContentControlDate(id, '2026-10-13');
  if (type.startsWith('checkbox')) return t.setContentControlChecked(id, false);
  return t.fillContentControl(id, '一般');
}

describe.each(['inline', 'block'])('相同值显式填写解除数据绑定：%s', (scope) => {
  it.each(cases)('$type 保留文字、格式与位置，只解除绑定，重复提交不记历史', ({ type, pr, content }) => {
    const editor = createTextEditor(parse(pr, content, scope));
    const before = editor.body;
    const span = contentControlSpans(before)[0];
    if (span === undefined) throw new Error('样本缺少控件');
    const id = span.control.id;
    const range = rangeOfContentControl(span);
    expect(span.control.dataBound).toBe(true);
    expect(
      editor.tx((t) => {
        expect(fill(t, type, id)).toEqual(range);
      }),
    ).toBeDefined();
    const after = editor.body;
    expect(after.sections).toEqual(before.sections);
    expect(contentControlSpans(after)[0]?.control).toEqual({ ...span.control, dataBound: undefined });
    expect(editor.tx((t) => void fill(t, type, id))).toBeUndefined();
    expect(editor.body).toBe(after);
    editor.undo();
    expect(editor.body).toBe(before);
    editor.redo();
    expect(editor.body).toBe(after);
  });
});

it.each(cases)('$type 仅解除绑定后发生错误仍全部回滚，不留下历史', ({ type, pr, content }) => {
  const editor = createTextEditor(parse(pr, content, 'inline'));
  const before = editor.body;
  const id = contentControlSpans(before)[0]?.control.id as string;
  expect(() =>
    editor.tx((t) => {
      fill(t, type, id);
      t.fillContentControl('missing', '失败');
    }),
  ).toThrow(/没有内容控件/);
  expect(editor.body).toBe(before);
  expect(editor.canUndo).toBe(false);
});

it.each(['dropDownList', 'comboBox'])('按序号重复选择 %s 当前项也解除绑定', (type) => {
  const sample = cases.find((c) => c.type === type);
  if (sample === undefined) throw new Error('样本缺少列表');
  const editor = createTextEditor(parse(sample.pr, sample.content, 'inline'));
  const before = editor.body;
  const id = contentControlSpans(before)[0]?.control.id as string;
  expect(editor.tx((t) => void t.selectContentControlItem(id, 0))).toBeDefined();
  expect(editor.body.sections).toEqual(before.sections);
  expect(contentControlSpans(editor.body)[0]?.control.dataBound).toBeUndefined();
  expect(editor.tx((t) => void t.selectContentControlItem(id, 0))).toBeUndefined();
});
