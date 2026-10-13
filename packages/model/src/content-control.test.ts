/**
 * 内容控件：解析出身份与成员，事务按 id 填值。
 * 回写（占位符标记 / 数据绑定怎么落回 XML）在 serialize 的 docx.test.ts。
 */
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, rangeOfContentControl, textOfContentControl } from './content-control.ts';
import type { Body, Paragraph } from './nodes.ts';
import { paragraphText, walkParagraphs } from './nodes.ts';
import { parseBody } from './parse-body.ts';
import { createTextEditor } from './text-transaction.ts';

function parse(xml: string): Body {
  return parseBody(parseXml(`<w:document><w:body>${xml}</w:body></w:document>`), createDiagnosticSink());
}

function paragraphs(body: Body): Paragraph[] {
  return [...walkParagraphs(body)];
}

function values(body: Body): Record<string, string> {
  return Object.fromEntries(contentControlSpans(body).map((s) => [s.control.tag, textOfContentControl(s)]));
}

const PLACEHOLDER =
  '<w:sdt><w:sdtPr><w:rPr><w:b/></w:rPr><w:alias w:val="申请人"/><w:tag w:val="applicant"/>' +
  '<w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>' +
  '<w:t>单击此处</w:t></w:r><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>输入文字。</w:t></w:r>' +
  '</w:sdtContent></w:sdt>';

describe('内容控件的解析', () => {
  it.each([
    ['', false],
    ['<w14:checked/>', true],
    ['<w14:checked w14:val="1"/>', true],
    ['<w14:checked w14:val="true"/>', true],
    ['<w14:checked w14:val="0"/>', false],
    ['<w14:checked w14:val="false"/>', false],
    ['<w14:checked w14:val="off"/>', false],
  ])('复选框状态按 w14:val 解析：%s → %s', (checked, expected) => {
    const body = parse(
      `<w:p><w:sdt><w:sdtPr><w14:checkbox>${checked}</w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
    );
    expect(contentControlSpans(body)[0]?.control.checkbox).toEqual({
      checked: expected,
      checkedState: { char: '☒', font: 'MS Gothic' },
      uncheckedState: { char: '☐', font: 'MS Gothic' },
    });
  });

  it('复选框自定义符号 / 字体保留，非法十六进制与代理码点回到默认符号', () => {
    const parseState = (states: string) =>
      contentControlSpans(
        parse(
          `<w:p><w:sdt><w:sdtPr><w14:checkbox>${states}</w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
        ),
      )[0]?.control.checkbox;
    expect(
      parseState(
        '<w14:checkedState w14:val="F0FE" w14:font="Wingdings"/><w14:uncheckedState w14:val="25a1" w14:font="宋体"/>',
      ),
    ).toMatchObject({
      checkedState: { char: '\uF0FE', font: 'Wingdings' },
      uncheckedState: { char: '□', font: '宋体' },
    });
    expect(
      parseState('<w14:checkedState w14:val="oops"/><w14:uncheckedState w14:val="D800"/>'),
    ).toMatchObject({
      checkedState: { char: '☒', font: 'MS Gothic' },
      uncheckedState: { char: '☐', font: 'MS Gothic' },
    });
  });

  it('行内控件：身份进控件表、run 带标记，块级控件标在段落上，嵌套靠 parent 串', () => {
    const body = parse(
      `<w:sdt><w:sdtPr><w:tag w:val="block"/><w:lock w:val="sdtLocked"/></w:sdtPr><w:sdtContent>` +
        `<w:p><w:r><w:t>甲</w:t></w:r>${PLACEHOLDER}<w:r><w:t>乙</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>丙</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>外</w:t></w:r></w:p>`,
    );
    const [p0, p1, p2] = paragraphs(body);
    const controls = Object.values(body.contentControls ?? {});
    const block = controls.find((c) => c.tag === 'block');
    const inline = controls.find((c) => c.tag === 'applicant');
    expect(block).toMatchObject({
      scope: 'block',
      type: 'richText',
      lock: 'sdtLocked',
      showingPlaceholder: false,
    });
    expect(inline).toMatchObject({
      scope: 'inline',
      type: 'text',
      alias: '申请人',
      parent: block?.id,
      showingPlaceholder: true,
      runProps: { bold: true },
    });
    expect(p0?.contentControl).toBe(block?.id);
    expect(p1?.contentControl).toBe(block?.id);
    expect(p2?.contentControl).toBeUndefined();
    expect(p0?.runs.map((r) => r.contentControl)).toEqual([undefined, inline?.id, inline?.id, undefined]);
    expect(values(body)).toEqual({ block: '甲单击此处输入文字。乙\n丙', applicant: '单击此处输入文字。' });
  });

  it('下拉框的选项、多行纯文本、数据绑定；没有控件的文档不带控件表', () => {
    const body = parse(
      '<w:p><w:sdt><w:sdtPr><w:tag w:val="level"/><w:dropDownList><w:listItem w:displayText="紧急" w:value="1"/>' +
        '<w:listItem w:value="一般"/></w:dropDownList></w:sdtPr><w:sdtContent><w:r><w:t>一般</w:t></w:r></w:sdtContent></w:sdt>' +
        '<w:sdt><w:sdtPr><w:tag w:val="memo"/><w:dataBinding w:xpath="/a"/><w:text w:multiLine="1"/></w:sdtPr>' +
        '<w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt></w:p>',
    );
    const [level, memo] = contentControlSpans(body).map((s) => s.control);
    expect(level?.items).toEqual([
      { text: '紧急', value: '1' },
      { text: '一般', value: '一般' },
    ]);
    expect(memo).toMatchObject({ type: 'text', multiLine: true, dataBound: true });
    expect(parse('<w:p/>').contentControls).toBeUndefined();
  });
});

describe('setContentControlDate', () => {
  const dateControl = (pr = '', content = '<w:r><w:t>某日</w:t></w:r>', extra = '') =>
    `<w:sdt><w:sdtPr><w:tag w:val="date"/>${extra}<w:date ${pr}><w:dateFormat w:val="yyyy年M月d日"/><w:lid w:val="zh-CN"/><w:calendar w:val="gregorian"/></w:date></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  const idOf = (body: Body) =>
    contentControlSpans(body).find((s) => s.control.tag === 'date')?.control.id as string;

  it('解析日期属性；填入日历日同步显示、元数据与占位格式，撤销重做同进同退', () => {
    const editor = createTextEditor(
      parse(
        `<w:p>${dateControl(
          'w:fullDate="2020-01-02T08:00:00+08:00"',
          '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>某日</w:t></w:r>',
          '<w:rPr><w:b/></w:rPr><w:showingPlcHdr/><w:dataBinding w:xpath="/date"/>',
        )}</w:p>`,
      ),
    );
    const before = editor.body;
    const id = idOf(before);
    expect(before.contentControls?.[id]?.date).toEqual({
      fullDate: '2020-01-02T08:00:00+08:00',
      format: 'yyyy年M月d日',
      language: 'zh-CN',
      calendar: 'gregorian',
    });
    let range: ReturnType<typeof rangeOfContentControl>;
    editor.tx((t) => {
      range = t.setContentControlDate(id, '2024-02-29');
    });
    const after = editor.body;
    const span = contentControlSpans(after)[0] as NonNullable<ReturnType<typeof contentControlSpans>[number]>;
    expect(values(after)).toEqual({ date: '2024年2月29日' });
    expect(span.runs).toHaveLength(1);
    expect(span.runs[0]?.props).toEqual({ bold: true });
    expect(span.control.date?.fullDate).toBe('2024-02-29T00:00:00Z');
    expect(span.control.showingPlaceholder).toBe(false);
    expect(span.control.dataBound).toBeUndefined();
    expect(range).toEqual(rangeOfContentControl(span));
    expect(
      editor.tx((t) => {
        t.setContentControlDate(id, '2024-02-29');
      }),
    ).toBeUndefined();
    expect(editor.body).toBe(after);
    editor.undo();
    expect(editor.body).toBe(before);
    editor.redo();
    expect(editor.body).toBe(after);
  });

  it('显示不含日时，只改 fullDate 也提交；自由文字填写清除旧日期，空块级控件可填', () => {
    const xml = dateControl('w:fullDate="2026-10-01T00:00:00Z"', '<w:r><w:t>2026年10月</w:t></w:r>').replace(
      'yyyy年M月d日',
      'yyyy年M月',
    );
    const editor = createTextEditor(parse(`<w:p>${xml}</w:p>`));
    const before = editor.body;
    const id = idOf(before);
    editor.tx((t) => {
      t.setContentControlDate(id, '2026-10-13');
    });
    expect(editor.body.sections).toEqual(before.sections);
    expect(editor.body.contentControls?.[id]?.date?.fullDate).toBe('2026-10-13T00:00:00Z');
    editor.tx((t) => {
      t.fillContentControl(id, '待定');
    });
    expect(values(editor.body)).toEqual({ date: '待定' });
    expect(editor.body.contentControls?.[id]?.date?.fullDate).toBeUndefined();
    editor.undo();
    expect(editor.body.contentControls?.[id]?.date?.fullDate).toBe('2026-10-13T00:00:00Z');
    const empty = createTextEditor(parse(dateControl('', '<w:p/>')));
    empty.tx((t) => {
      t.setContentControlDate(idOf(empty.body), '0001-01-01');
    });
    expect(values(empty.body)).toEqual({ date: '0001年1月1日' });
  });

  it.each([
    '2026-02-29',
    '2024-02-30',
    '2026-00-01',
    '2026-13-01',
    '0000-01-01',
    '2026-1-1',
    '2026-10-13T00:00:00Z',
  ])('拒绝非法日历日 %s，先前文字修改也原子回滚', (value) => {
    const editor = createTextEditor(parse(`<w:p><w:r><w:t>外</w:t></w:r>${dateControl()}</w:p>`));
    const before = editor.body;
    expect(() =>
      editor.tx((t) => {
        t.insertText(
          { nodeId: paragraphs(before)[0]?.runs[0]?.id as string, contentIndex: 0, offset: 1 },
          '改',
        );
        t.setContentControlDate(idOf(before), value);
      }),
    ).toThrow(/日期/);
    expect(editor.body).toBe(before);
    expect(editor.canUndo).toBe(false);
  });

  it('内容锁沿父链生效；修订、域、多段、多 run 与嵌套内容拒绝', () => {
    for (const content of [
      '<w:ins><w:r><w:t>某日</w:t></w:r></w:ins>',
      '<w:fldSimple w:instr="DATE"><w:r><w:t>某日</w:t></w:r></w:fldSimple>',
      '<w:p><w:r><w:t>某</w:t></w:r></w:p><w:p><w:r><w:t>日</w:t></w:r></w:p>',
      '<w:r><w:t>某</w:t></w:r><w:r><w:t>日</w:t></w:r>',
      '<w:sdt><w:sdtContent><w:r><w:t>某日</w:t></w:r></w:sdtContent></w:sdt>',
    ]) {
      const editor = createTextEditor(
        parse(
          content.startsWith('<w:p>') ? dateControl('', content) : `<w:p>${dateControl('', content)}</w:p>`,
        ),
      );
      const before = editor.body;
      expect(() =>
        editor.tx((t) => {
          t.setContentControlDate(idOf(before), '2026-10-13');
        }),
      ).toThrow();
      expect(editor.body).toBe(before);
    }
    const editor = createTextEditor(
      parse(
        `<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:p>${dateControl()}</w:p></w:sdtContent></w:sdt>`,
      ),
    );
    const before = editor.body;
    expect(() =>
      editor.tx((t) => {
        t.setContentControlDate(idOf(before), '2026-10-13');
      }),
    ).toThrow(/锁定/);
    expect(editor.body).toBe(before);
  });

  it.each([
    ['w:val="gregorian"', 'w:val="hijri"', /历法/],
    ['w:val="zh-CN"', 'w:val="fr-FR"', /语言/],
    ['<w:dateFormat w:val="yyyy年M月d日"/>', '', /dateFormat/],
    ['<w:date ', '<w:text ', /不是日期/],
  ])('不支持的日期设置明确拒绝', (from, to, error) => {
    const editor = createTextEditor(parse(`<w:p>${dateControl().replace(from, to)}</w:p>`));
    expect(() =>
      editor.tx((t) => {
        t.setContentControlDate(idOf(editor.body), '2026-10-13');
      }),
    ).toThrow(error);
  });
});

describe('setContentControlChecked', () => {
  const checkbox = (pr = '', content = '<w:r><w:rPr><w:b/></w:rPr><w:t>☐</w:t></w:r>', extra = '') =>
    `<w:sdt><w:sdtPr><w:tag w:val="agree"/>${extra}<w14:checkbox>${pr}</w14:checkbox></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  const idOf = (body: Body) =>
    contentControlSpans(body).find((s) => s.control.tag === 'agree')?.control.id as string;
  const checkedOf = (body: Body) => body.contentControls?.[idOf(body)]?.checkbox?.checked;

  it('切换状态与符号是同一撤销单元，保留 run 格式与位置；重复设置无修改', () => {
    const editor = createTextEditor(parse(`<w:p><w:r><w:t>外</w:t></w:r>${checkbox()}</w:p>`));
    const before = editor.body;
    const id = idOf(before);
    let range: ReturnType<typeof rangeOfContentControl>;
    const change = editor.tx((t) => {
      range = t.setContentControlChecked(id, true);
    });
    expect(change).toBeDefined();
    expect(values(editor.body)).toEqual({ agree: '☒' });
    expect(checkedOf(editor.body)).toBe(true);
    const span = contentControlSpans(editor.body)[0];
    expect(span?.runs[0]).toMatchObject({
      id: contentControlSpans(before)[0]?.runs[0]?.id,
      props: { bold: true },
      content: [{ kind: 'symbol', char: '☒', font: 'MS Gothic' }],
    });
    expect(range).toEqual(rangeOfContentControl(span as NonNullable<typeof span>));
    const after = editor.body;
    expect(
      editor.tx((t) => {
        t.setContentControlChecked(id, true);
      }),
    ).toBeUndefined();
    expect(editor.body).toBe(after);
    editor.undo();
    expect(editor.body).toBe(before);
    editor.redo();
    expect(editor.body).toBe(after);
    editor.tx((t) => {
      t.setContentControlChecked(id, false);
    });
    expect(values(editor.body)).toEqual({ agree: '☐' });
    expect(checkedOf(editor.body)).toBe(false);
  });

  it('自定义符号用模板指定的字体，占位符样式换成控件格式，填写后解除数据绑定', () => {
    const editor = createTextEditor(
      parse(
        `<w:p>${checkbox(
          '<w14:checkedState w14:val="F0FE" w14:font="Wingdings"/>',
          '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>☐</w:t></w:r>',
          '<w:rPr><w:i/></w:rPr><w:showingPlcHdr/><w:dataBinding w:xpath="/agree"/>',
        )}</w:p>`,
      ),
    );
    const id = idOf(editor.body);
    editor.tx((t) => {
      t.setContentControlChecked(id, true);
    });
    const span = contentControlSpans(editor.body)[0];
    expect(span?.runs[0]?.content).toEqual([{ kind: 'symbol', char: '\uF0FE', font: 'Wingdings' }]);
    expect(span?.runs[0]?.props).toEqual({ italic: true });
    expect(span?.control).toMatchObject({ showingPlaceholder: false, checkbox: { checked: true } });
    expect(span?.control.dataBound).toBeUndefined();
  });

  it('两种状态使用同一个符号时，仅状态改变也记录历史，不复制字符；无修改保持原树', () => {
    const editor = createTextEditor(
      parse(
        `<w:p>${checkbox(
          '<w14:checkedState w14:val="2610" w14:font="宋体"/><w14:uncheckedState w14:val="2610" w14:font="宋体"/>',
          '<w:r><w:sym w:font="宋体" w:char="2610"/></w:r>',
        )}</w:p>`,
      ),
    );
    const before = editor.body;
    const id = idOf(before);
    expect(
      editor.tx((t) => {
        t.setContentControlChecked(id, false);
      }),
    ).toBeUndefined();
    expect(editor.body).toBe(before);
    expect(
      editor.tx((t) => {
        t.setContentControlChecked(id, true);
      }),
    ).toBeDefined();
    expect(editor.body.sections).toEqual(before.sections);
    expect(checkedOf(editor.body)).toBe(true);
    editor.undo();
    expect(editor.body).toBe(before);
  });

  it.each(['contentLocked', 'sdtContentLocked'])('自身或外层 %s 锁拒绝，sdtLocked 仍可填写', (lock) => {
    for (const xml of [
      `<w:p>${checkbox('', undefined, `<w:lock w:val="${lock}"/>`)}</w:p>`,
      `<w:sdt><w:sdtPr><w:lock w:val="${lock}"/></w:sdtPr><w:sdtContent><w:p>${checkbox()}</w:p></w:sdtContent></w:sdt>`,
    ]) {
      const editor = createTextEditor(parse(xml));
      const before = editor.body;
      expect(() =>
        editor.tx((t) => {
          t.setContentControlChecked(idOf(before), true);
        }),
      ).toThrow(/锁定/);
      expect(editor.body).toBe(before);
    }
    const unlocked = createTextEditor(
      parse(`<w:p>${checkbox('', undefined, '<w:lock w:val="sdtLocked"/>')}</w:p>`),
    );
    unlocked.tx((t) => {
      t.setContentControlChecked(idOf(unlocked.body), true);
    });
    expect(checkedOf(unlocked.body)).toBe(true);
  });

  it('拒绝多字符、修订、域及错误参数，先前命令也原子回滚', () => {
    for (const content of [
      '<w:r><w:t>多字</w:t></w:r>',
      '<w:ins w:id="1" w:author="张三"><w:r><w:t>☐</w:t></w:r></w:ins>',
      '<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>',
    ]) {
      const editor = createTextEditor(parse(`<w:p><w:r><w:t>外</w:t></w:r>${checkbox('', content)}</w:p>`));
      const before = editor.body;
      expect(() =>
        editor.tx((t) => {
          t.insertText(
            { nodeId: paragraphs(before)[0]?.runs[0]?.id as string, contentIndex: 0, offset: 1 },
            '改',
          );
          t.setContentControlChecked(idOf(before), true);
        }),
      ).toThrow(/单个无修订/);
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(false);
    }
    const editor = createTextEditor(parse(`<w:p>${checkbox()}</w:p>`));
    expect(() =>
      editor.tx((t) => {
        t.setContentControlChecked(idOf(editor.body), 'true' as unknown as boolean);
      }),
    ).toThrow(/布尔/);
    expect(() =>
      editor.tx((t) => {
        t.setContentControlChecked('missing', true);
      }),
    ).toThrow(/没有内容控件/);
    const text = createTextEditor(parse(`<w:p>${PLACEHOLDER}</w:p>`));
    expect(() =>
      text.tx((t) => {
        t.setContentControlChecked(contentControlSpans(text.body)[0]?.control.id as string, true);
      }),
    ).toThrow(/不是复选框/);
  });
});

describe('fillContentControl', () => {
  it('内层文字控件受外层内容锁约束，失败时不产生撤销单元', () => {
    const editor = createTextEditor(
      parse(
        `<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:p>${PLACEHOLDER}</w:p></w:sdtContent></w:sdt>`,
      ),
    );
    const before = editor.body;
    const id = contentControlSpans(before).find((s) => s.control.tag === 'applicant')?.control.id as string;
    expect(() =>
      editor.tx((t) => {
        t.fillContentControl(id, '张三');
      }),
    ).toThrow(/锁定/);
    expect(editor.body).toBe(before);
    expect(editor.canUndo).toBe(false);
  });

  it('占位符控件：内容换成值、换上控件自己的格式、清掉占位标记，一次撤销全回去', () => {
    const editor = createTextEditor(parse(`<w:p><w:r><w:t>申请人：</w:t></w:r>${PLACEHOLDER}</w:p>`));
    const before = editor.body;
    const id = contentControlSpans(before)[0]?.control.id as string;
    let filled: ReturnType<typeof rangeOfContentControl>;
    editor.tx((t) => {
      filled = t.fillContentControl(id, '张三');
    });
    const p = paragraphs(editor.body)[0] as Paragraph;
    expect(paragraphText(p)).toBe('申请人：张三');
    const run = p.runs.find((r) => r.content.some((c) => c.kind === 'text' && c.text === '张三'));
    expect(run?.contentControl).toBe(id);
    expect(run?.props).toEqual({ bold: true });
    expect(editor.body.contentControls?.[id]?.showingPlaceholder).toBe(false);
    expect(values(editor.body)).toEqual({ applicant: '张三' });
    expect(filled?.start.nodeId).toBe(run?.id);
    editor.undo();
    expect(editor.body).toBe(before);
    expect(before.contentControls?.[id]?.showingPlaceholder).toBe(true);
  });

  it('换行与制表位写成软换行 / 制表位；块级控件跨段的内容并成一段', () => {
    const editor = createTextEditor(
      parse(
        '<w:sdt><w:sdtPr><w:tag w:val="body"/></w:sdtPr><w:sdtContent>' +
          '<w:p><w:r><w:t>旧一</w:t></w:r></w:p><w:p><w:r><w:t>旧二</w:t></w:r></w:p>' +
          '</w:sdtContent></w:sdt><w:p><w:r><w:t>尾</w:t></w:r></w:p>',
      ),
    );
    const id = contentControlSpans(editor.body)[0]?.control.id as string;
    editor.tx((t) => {
      t.fillContentControl(id, '第一行\n甲\t乙');
    });
    const ps = paragraphs(editor.body);
    expect(ps.map(paragraphText)).toEqual(['第一行\n甲\t乙', '尾']);
    expect(ps[0]?.contentControl).toBe(id);
    expect(values(editor.body)).toEqual({ body: '第一行\n甲\t乙' });
  });

  it('下拉框认显示文字或值，不在选项里拒绝；纯文本单行控件把换行换成空格', () => {
    const editor = createTextEditor(
      parse(
        '<w:p><w:sdt><w:sdtPr><w:tag w:val="level"/><w:dropDownList><w:listItem w:displayText="紧急" w:value="1"/>' +
          '</w:dropDownList></w:sdtPr><w:sdtContent><w:r><w:t>选择</w:t></w:r></w:sdtContent></w:sdt>' +
          '<w:sdt><w:sdtPr><w:tag w:val="name"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>名</w:t></w:r></w:sdtContent></w:sdt></w:p>',
      ),
    );
    const [level, name] = contentControlSpans(editor.body).map((s) => s.control.id as string);
    editor.tx((t) => {
      t.fillContentControl(level as string, '1');
      t.fillContentControl(name as string, '甲\n乙');
    });
    expect(values(editor.body)).toEqual({ level: '紧急', name: '甲 乙' });
    const after = editor.body;
    expect(() =>
      editor.tx((t) => {
        t.fillContentControl(level as string, '不存在');
      }),
    ).toThrow(/下拉框没有这个选项/);
    expect(editor.body).toBe(after);
  });

  it('锁了内容 / 复选框 / 不存在的控件拒绝；填外层控件时内层控件随内容一起消失', () => {
    const editor = createTextEditor(
      parse(
        '<w:p><w:sdt><w:sdtPr><w:tag w:val="locked"/><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>' +
          '<w:r><w:t>锁</w:t></w:r></w:sdtContent></w:sdt>' +
          '<w:sdt><w:sdtPr><w:tag w:val="check"/><w14:checkbox/></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>' +
          '<w:sdt><w:sdtPr><w:tag w:val="outer"/></w:sdtPr><w:sdtContent><w:r><w:t>前</w:t></w:r>' +
          '<w:sdt><w:sdtPr><w:tag w:val="inner"/></w:sdtPr><w:sdtContent><w:r><w:t>里</w:t></w:r></w:sdtContent></w:sdt>' +
          '</w:sdtContent></w:sdt></w:p>',
      ),
    );
    const ids = Object.fromEntries(
      contentControlSpans(editor.body).map((s) => [s.control.tag, s.control.id]),
    );
    expect(() => editor.tx((t) => void t.fillContentControl(ids.locked as string, 'x'))).toThrow(/锁定/);
    expect(() => editor.tx((t) => void t.fillContentControl(ids.check as string, 'x'))).toThrow(/checkbox/);
    expect(() => editor.tx((t) => void t.fillContentControl('nope', 'x'))).toThrow(/没有内容控件/);
    expect(values(editor.body)).toMatchObject({ outer: '前里', inner: '里' });
    editor.tx((t) => {
      t.fillContentControl(ids.outer as string, '新');
    });
    const tags = contentControlSpans(editor.body).map((s) => s.control.tag);
    expect(tags).toEqual(['locked', 'check', 'outer']);
    expect(values(editor.body)).toMatchObject({ outer: '新' });
  });
});
