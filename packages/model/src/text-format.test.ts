import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { Body, Paragraph } from './nodes.ts';
import { paragraphText, walkParagraphs } from './nodes.ts';
import { rangeOfNode } from './order.ts';
import { parseBody } from './parse-body.ts';
import type { DocPosition, DocRange } from './position.ts';
import { textOfRange } from './range-text.ts';
import { mapTextPosition, mapTextRange } from './text-change.ts';
import { createTextEditor } from './text-transaction.ts';

function parse(xml: string): Body {
  return parseBody(parseXml(`<w:document><w:body>${xml}</w:body></w:document>`), createDiagnosticSink());
}

function para(body: Body, index = 0): Paragraph {
  return [...walkParagraphs(body)][index] as Paragraph;
}

function pos(body: Body, offset = 0, run = 0, contentIndex = 0, paragraph = 0): DocPosition {
  return { nodeId: para(body, paragraph).runs[run]?.id as string, contentIndex, offset };
}

/** 直接格式树即可复用复制逻辑：只读 hidden 与片段文字。 */
function textOf(body: Body, range: DocRange): string {
  return textOfRange(body as never, range);
}

describe('字符格式命令', () => {
  it('范围在 run 中间时拆成三段，只改中间一段并返回拆分后的同一段文字', () => {
    const editor = createTextEditor(
      parse(
        '<w:p><w:hyperlink r:id="rId1"><w:r><w:rPr><w:i/></w:rPr><w:t>甲乙丙丁</w:t></w:r></w:hyperlink></w:p>',
      ),
    );
    const before = editor.body;
    const range = { start: pos(before, 1), end: pos(before, 3) };
    let result: DocRange | undefined;
    const changes = editor.tx((t) => {
      result = t.setRunProps(range, { bold: true });
    });
    const runs = para(editor.body).runs;
    expect(runs.map((r) => paragraphText({ ...para(editor.body), runs: [r] }))).toEqual(['甲', '乙丙', '丁']);
    expect(runs.map((r) => r.props.bold)).toEqual([undefined, true, undefined]);
    expect(runs.every((r) => r.props.italic && r.hyperlink?.relId === 'rId1')).toBe(true);
    expect(runs[0]?.id).toBe(para(before).runs[0]?.id);
    expect(new Set(runs.map((r) => r.id)).size).toBe(3);
    expect(textOf(editor.body, result as DocRange)).toBe('乙丙');
    // 选区与光标经变更集映射后仍指着同一段文字。
    expect(textOf(editor.body, mapTextRange(range, changes as never))).toBe('乙丙');
    expect(mapTextPosition(pos(before, 4), changes as never)).toEqual({
      nodeId: runs[2]?.id,
      contentIndex: 0,
      offset: 1,
    });
    expect(changes?.paragraphIds).toEqual([para(before).id]);
    editor.undo();
    expect(editor.body).toBe(before);
    editor.redo();
    expect(para(editor.body).runs[1]?.props.bold).toBe(true);
  });

  it('跨段落与表格修改，整段覆盖不拆 run，越过段尾时一并修改段落标记', () => {
    const editor = createTextEditor(
      parse(
        '<w:p><w:r><w:t>前段</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>格内</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/><w:p><w:r><w:t>末段文字</w:t></w:r></w:p>',
      ),
    );
    const before = editor.body;
    const ids = para(before).runs.map((r) => r.id);
    editor.tx((t) => {
      t.setRunProps({ start: pos(before), end: pos(before, 2, 0, 0, 3) }, { underline: 'single' });
    });
    const body = editor.body;
    expect(para(body).runs.map((r) => r.id)).toEqual(ids);
    expect(para(body).runs[0]?.props.underline).toBe('single');
    expect(para(body, 1).runs[0]?.props.underline).toBe('single');
    expect(para(body, 3).runs.map((r) => r.props.underline)).toEqual(['single', undefined]);
    // 前三段的段落标记被范围越过；末段标记没选中，编号与空段输入不应跟着变。
    expect([0, 1, 2, 3].map((i) => para(body, i).props.markRunProps?.underline)).toEqual([
      'single',
      'single',
      'single',
      undefined,
    ]);
  });

  it('折叠在空段落时修改段落标记，首次输入即带上新格式；非空段落内折叠是空操作', () => {
    const editor = createTextEditor(parse('<w:p/><w:p><w:r><w:t>正文</w:t></w:r></w:p>'));
    const empty = { nodeId: para(editor.body).id, contentIndex: 0, offset: 0 };
    editor.tx((t) => {
      t.setRunProps({ start: empty, end: empty }, { bold: true });
    });
    expect(para(editor.body).props.markRunProps?.bold).toBe(true);
    editor.tx((t) => {
      t.insertText(empty, '新');
    });
    expect(para(editor.body).runs[0]?.props.bold).toBe(true);
    const at = pos(editor.body, 1, 0, 0, 1);
    const unchanged = editor.body;
    expect(
      editor.tx((t) => {
        t.setRunProps({ start: at, end: at }, { bold: true });
      }),
    ).toBeUndefined();
    expect(editor.body).toBe(unchanged);
  });

  it('null 删除直接格式回到样式值，值未变化时不拆 run 也不产生撤销单元', () => {
    const editor = createTextEditor(
      parse('<w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>加粗文字</w:t></w:r></w:p>'),
    );
    const before = editor.body;
    const range = { start: pos(before, 1), end: pos(before, 3) };
    expect(
      editor.tx((t) => {
        t.setRunProps(range, { bold: true });
      }),
    ).toBeUndefined();
    expect(editor.canUndo).toBe(false);
    editor.tx((t) => {
      t.setRunProps(rangeOfNode(para(before)) as DocRange, { bold: null });
    });
    expect(para(editor.body).runs).toHaveLength(1);
    expect(para(editor.body).runs[0]?.props).toEqual({ size: 280 });
  });

  it('端点落在片段边界时不留空文字片段，后方片段下标随新 run 重排', () => {
    const editor = createTextEditor(parse('<w:p><w:r><w:t>甲乙</w:t><w:tab/><w:t>丙丁</w:t></w:r></w:p>'));
    const before = editor.body;
    const changes = editor.tx((t) => {
      t.setRunProps({ start: pos(before, 2), end: pos(before, 0, 0, 2) }, { italic: true });
    });
    const runs = para(editor.body).runs;
    expect(runs.map((r) => r.content)).toEqual([
      [{ kind: 'text', text: '甲乙' }],
      [{ kind: 'tab' }],
      [{ kind: 'text', text: '丙丁' }],
    ]);
    expect(runs.map((r) => r.props.italic)).toEqual([undefined, true, undefined]);
    expect(mapTextPosition(pos(before, 1, 0, 2), changes as never)).toEqual({
      nodeId: runs[2]?.id,
      contentIndex: 0,
      offset: 1,
    });
  });

  it('非法范围与失败命令整批回滚', () => {
    const editor = createTextEditor(parse('<w:p><w:r><w:t>甲乙</w:t></w:r></w:p>'));
    const before = editor.body;
    expect(() =>
      editor.tx((t) => {
        t.setRunProps({ start: pos(before, 2), end: pos(before, 1) }, { bold: true });
      }),
    ).toThrow('起点');
    expect(() =>
      editor.tx((t) => {
        t.setRunProps({ start: pos(before), end: pos(before, 9) }, { bold: true });
      }),
    ).toThrow();
    expect(() =>
      editor.tx((t) => {
        t.setRunProps({ start: pos(before), end: pos(before, 1) }, { bold: true });
        t.insertText(pos(before, 99), '错');
      }),
    ).toThrow();
    expect(editor.body).toBe(before);
  });
});
