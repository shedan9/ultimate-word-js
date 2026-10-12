/**
 * 批注的编辑：新建 / 回复 / 改正文 / 解决 / 删除都是事务命令，内容与 run 上的标记同进同退、随撤销复原。
 * 判据是 `listComments` 答出来的东西（范围、引文、线程），不看节点长什么样。
 */
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { commentIdSeed } from './comment-edit.ts';
import { listComments } from './comments.ts';
import type { Body, Paragraph } from './nodes.ts';
import { paragraphText, walkParagraphs } from './nodes.ts';
import { parseBody } from './parse-body.ts';
import type { DocPosition } from './position.ts';
import { createTextEditor } from './text-transaction.ts';

function parse(xml: string): Body {
  return parseBody(parseXml(`<w:document><w:body>${xml}</w:body></w:document>`), createDiagnosticSink());
}

const TWO = '<w:p><w:r><w:t>甲乙丙丁</w:t></w:r></w:p><w:p><w:r><w:t>戊己</w:t></w:r></w:p>';

function para(body: Body, index = 0): Paragraph {
  return [...walkParagraphs(body)][index] as Paragraph;
}

function at(body: Body, paragraph: number, offset: number): DocPosition {
  return { nodeId: para(body, paragraph).runs[0]?.id as string, contentIndex: 0, offset };
}

describe('批注的编辑', () => {
  it('新建：范围里的字打上标记（端点拆 run，可跨段），引文与作者对得上；撤销连内容一起回去', () => {
    const editor = createTextEditor(parse(TWO));
    const before = editor.body;
    let id = '';
    editor.tx((t) => {
      id = t.addComment(
        { start: at(before, 0, 2), end: at(before, 1, 1) },
        { text: '改一下\n第二段', author: '张三', initials: '张', date: '2026-10-12T09:00:00Z' },
      );
    });
    expect(id).toBe('0');
    const [c] = listComments(editor.body);
    expect(c).toMatchObject({
      id: '0',
      author: '张三',
      initials: '张',
      text: '改一下\n第二段',
      quote: '丙丁\n戊',
    });
    expect(c?.done).toBe(false);
    // 首段拆成「甲乙 | 丙丁」，次段拆成「戊 | 己」，文字不变
    expect(para(editor.body).runs.map((r) => r.comments ?? [])).toEqual([[], ['0']]);
    expect(paragraphText(para(editor.body, 1))).toBe('戊己');
    editor.undo();
    expect(editor.body).toBe(before);
    expect(listComments(editor.body)).toEqual([]);
  });

  it('折叠范围是点批注：插一个空 run 当锚，range 是折叠的、落在旁边的字上', () => {
    const editor = createTextEditor(parse(TWO));
    const body = editor.body;
    editor.tx((t) => {
      t.addComment({ start: at(body, 0, 2), end: at(body, 0, 2) }, { text: '这里' });
    });
    const runs = para(editor.body).runs;
    expect(runs.map((r) => [r.content.length, r.comments ?? []])).toEqual([
      [1, []],
      [0, ['0']],
      [1, []],
    ]);
    const [c] = listComments(editor.body);
    expect(c?.quote).toBe('');
    expect(c?.range?.start).toEqual(c?.range?.end);
    expect(paragraphText(para(editor.body))).toBe('甲乙丙丁');
  });

  it('回复罩着与顶层相同的字；回复的回复挂到顶层上；删顶层连回复一起删', () => {
    const editor = createTextEditor(parse(TWO));
    const body = editor.body;
    let root = '';
    let reply = '';
    let nested = '';
    editor.tx((t) => {
      root = t.addComment({ start: at(body, 0, 0), end: at(body, 0, 2) }, { text: '一' });
    });
    editor.tx((t) => {
      reply = t.replyComment(root, { text: '二', author: '李四' });
      nested = t.replyComment(reply, { text: '三' });
    });
    const list = listComments(editor.body);
    expect(list.map((c) => [c.id, c.parent, c.quote])).toEqual([
      [root, undefined, '甲乙'],
      [reply, root, '甲乙'],
      [nested, root, '甲乙'],
    ]);
    editor.tx((t) => {
      t.deleteComment(reply);
    });
    expect(listComments(editor.body).map((c) => c.id)).toEqual([root, nested]);
    editor.tx((t) => {
      t.deleteComment(root);
    });
    expect(listComments(editor.body)).toEqual([]);
    expect(editor.body.comments).toBeUndefined();
    expect(para(editor.body).runs.every((r) => r.comments === undefined)).toBe(true);
  });

  it('改正文与解决：解决的是整条线程；没变化时整次无修改', () => {
    const editor = createTextEditor(parse(TWO));
    const body = editor.body;
    let root = '';
    let reply = '';
    editor.tx((t) => {
      root = t.addComment({ start: at(body, 0, 0), end: at(body, 0, 1) }, { text: '旧' });
      reply = t.replyComment(root, { text: '回' });
    });
    expect(
      editor.tx((t) => {
        t.setCommentText(root, '旧');
      }),
    ).toBeUndefined();
    expect(
      editor.tx((t) => {
        t.setCommentText(root, '新\t字');
      }),
    ).toBeDefined();
    expect(listComments(editor.body)[0]?.text).toBe('新\t字');
    editor.tx((t) => {
      t.setCommentDone(reply, true);
    });
    expect(listComments(editor.body).map((c) => c.done)).toEqual([true, true]);
    expect(
      editor.tx((t) => {
        t.setCommentDone(root, true);
      }),
    ).toBeUndefined();
    editor.undo();
    expect(listComments(editor.body).map((c) => c.done)).toEqual([false, false]);
  });

  it('新 id 接着用过的最大号、不复用删掉的（回写按 id 认原文里那条）', () => {
    const source = parse(
      '<w:p><w:commentRangeStart w:id="7"/><w:r><w:t>甲</w:t></w:r><w:commentRangeEnd w:id="7"/></w:p>',
    );
    expect(commentIdSeed(source)).toBe(8);
    const editor = createTextEditor(source);
    const body = editor.body;
    const range = { start: at(body, 0, 0), end: at(body, 0, 1) };
    let a = '';
    editor.tx((t) => {
      a = t.addComment(range, { text: 'a' });
    });
    editor.undo();
    let b = '';
    editor.tx((t) => {
      b = t.addComment(range, { text: 'b' });
    });
    expect([a, b]).toEqual(['8', '9']);
  });

  it('批注不存在、正文不是字符串、范围反了都整次拒绝', () => {
    const editor = createTextEditor(parse(TWO));
    const body = editor.body;
    expect(() =>
      editor.tx((t) => {
        t.deleteComment('3');
      }),
    ).toThrow('没有批注');
    expect(() => editor.tx((t) => void t.replyComment('3', { text: 'x' }))).toThrow('没有批注');
    expect(() =>
      editor.tx(
        (t) => void t.addComment({ start: at(body, 0, 0), end: at(body, 0, 0) }, { text: 1 as never }),
      ),
    ).toThrow('字符串');
    expect(() =>
      editor.tx((t) => void t.addComment({ start: at(body, 0, 2), end: at(body, 0, 1) }, { text: 'x' })),
    ).toThrow('起点');
    expect(editor.canUndo).toBe(false);
  });
});
