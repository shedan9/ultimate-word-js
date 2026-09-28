/**
 * 内容控件：解析出身份与成员，事务按 id 填值。
 * 回写（占位符标记 / 数据绑定怎么落回 XML）在 serialize 的 docx.test.ts。
 */
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { contentControlSpans, type rangeOfContentControl, textOfContentControl } from './content-control.ts';
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

describe('fillContentControl', () => {
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
