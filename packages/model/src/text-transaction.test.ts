import { readFileSync } from 'node:fs';
import { createDiagnosticSink } from '@uw/core';
import { OpcPackage, parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { loadDocument } from './load.ts';
import type { Body, Paragraph } from './nodes.ts';
import { paragraphText, walkParagraphs } from './nodes.ts';
import { parseBody } from './parse-body.ts';
import type { DocPosition, DocRange } from './position.ts';
import { resolveBody } from './resolve-body.ts';
import { findText } from './search.ts';
import type { TextChangeSet } from './text-change.ts';
import { mapTextPosition, mapTextRange } from './text-change.ts';
import type { TextEditor, TextTransaction } from './text-transaction.ts';
import { createTextEditor } from './text-transaction.ts';

function must<T>(value: T | undefined): T {
  if (value === undefined) throw new Error('测试缺少预期结果');
  return value;
}

function parse(xml = '<w:p><w:r><w:t>中文正文</w:t></w:r></w:p>'): Body {
  return parseBody(parseXml(`<w:document><w:body>${xml}</w:body></w:document>`), createDiagnosticSink());
}

function para(body: Body, index = 0): Paragraph {
  return [...walkParagraphs(body)][index] as Paragraph;
}

function pos(body: Body, offset = 0, run = 0, contentIndex = 0, paragraph = 0): DocPosition {
  return { nodeId: para(body, paragraph).runs[run]?.id as string, contentIndex, offset };
}

function insert(editor: TextEditor, at: DocPosition, text: string, input = false): TextChangeSet {
  return editor.tx(
    (t) => {
      t.insertText(at, text);
    },
    { origin: input ? 'input' : 'command' },
  ) as TextChangeSet;
}

describe('文字事务', () => {
  it('一批命令原子提交，读取侧只见提交前后的快照，撤销恢复完整原树', () => {
    const source = parse();
    const editor = createTextEditor(source);
    const before = editor.body;
    const at = pos(before, 2);
    const result = editor.tx((t) => {
      const caret = t.insertText(at, '新增');
      t.insertText(caret, '内容');
      expect(editor.body).toBe(before);
    });
    expect(paragraphText(para(editor.body))).toBe('中文新增内容正文');
    expect(result?.paragraphIds).toEqual([para(before).id]);
    expect(result?.changes).toHaveLength(2);
    expect(editor.undo()?.changes.map((c) => c.deletedText)).toEqual(['内容', '新增']);
    expect(editor.body).toBe(before);
    expect(editor.canUndo).toBe(false);
    expect(editor.canRedo).toBe(true);
    editor.redo();
    expect(paragraphText(para(editor.body))).toBe('中文新增内容正文');
    expect(source).toEqual(before);
    expect(structuredClone(editor.body)).toEqual(editor.body);
  });

  it('调用方改源树或修改返回的变更集，不会污染快照与撤销记录', () => {
    const source = parse();
    const editor = createTextEditor(source);
    para(source).runs.length = 0;
    const result = insert(editor, pos(editor.body), '前');
    must(result.changes[0]).insertedText = '错误';
    result.changes.length = 0;
    expect(() => {
      para(editor.body).runs.length = 0;
    }).toThrow();
    expect(editor.undo()?.changes[0]?.deletedText).toBe('前');
    expect(paragraphText(para(editor.body))).toBe('中文正文');
  });

  it('跨 run 与片段删除保留样式、链接、id 和后方片段下标', () => {
    const editor = createTextEditor(
      parse(
        '<w:p><w:r><w:t>甲乙</w:t><w:t>丙丁</w:t></w:r><w:hyperlink r:id="rId1"><w:r><w:rPr><w:b/></w:rPr><w:t>戊己</w:t><w:t>庚辛</w:t></w:r></w:hyperlink></w:p>',
      ),
    );
    const before = editor.body;
    const start = pos(before, 1);
    const end = pos(before, 1, 1);
    const changes = editor.tx((t) => {
      t.deleteRange({ start, end });
    }) as TextChangeSet;
    expect(paragraphText(para(editor.body))).toBe('甲己庚辛');
    expect(para(editor.body).runs.map((r) => r.id)).toEqual(para(before).runs.map((r) => r.id));
    expect(para(editor.body).runs[0]?.content[1]).toEqual({ kind: 'text', text: '' });
    expect(para(editor.body).runs[1]?.props.bold).toBe(true);
    expect(para(editor.body).runs[1]?.hyperlink).toEqual({ relId: 'rId1' });
    expect(mapTextPosition(pos(before, 2, 1), changes)).toEqual(pos(before, 1, 1));
    expect(mapTextPosition(pos(before, 1, 1, 1), changes)).toEqual(pos(before, 1, 1, 1));
    editor.undo();
    expect(editor.body).toEqual(before);
  });

  it('已存在的空 run 可以输入，撤销恢复没有片段的原样', () => {
    const editor = createTextEditor(parse('<w:p><w:r/></w:p>'));
    insert(editor, pos(editor.body), '空');
    expect(paragraphText(para(editor.body))).toBe('空');
    editor.undo();
    expect(para(editor.body).runs[0]?.content).toEqual([]);
  });

  it('嵌套表格内文字可改，未修改的相邻段落保留引用', () => {
    const editor = createTextEditor(
      parse(
        '<w:p><w:r><w:t>外</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:tbl><w:tr><w:tc><w:p><w:r><w:t>内</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>',
      ),
    );
    const outside = para(editor.body);
    insert(editor, pos(editor.body, 1, 0, 0, 1), '容');
    expect(para(editor.body)).toBe(outside);
    expect(paragraphText(para(editor.body, 1))).toBe('内容');
  });

  it('命令错误即使被回调捕获也回滚整批，原有 redo 不丢', () => {
    const editor = createTextEditor(parse());
    insert(editor, pos(editor.body), '前');
    editor.undo();
    const before = editor.body;
    expect(() =>
      editor.tx((t) => {
        t.insertText(pos(before), '新');
        expect(() => t.insertText(pos(before, 999), '错')).toThrow();
      }),
    ).toThrow('全部回滚');
    expect(editor.body).toBe(before);
    expect(editor.canUndo).toBe(false);
    expect(editor.canRedo).toBe(true);
  });

  it('回调抛错、重入事务、操作历史都不能泄漏半次修改', () => {
    const editor = createTextEditor(parse());
    const before = editor.body;
    for (const action of [
      () => {
        throw new Error('回调失败');
      },
      () => editor.tx(() => {}),
      () => editor.undo(),
      () => editor.redo(),
      () => editor.breakHistory(),
    ]) {
      expect(() =>
        editor.tx((t) => {
          t.insertText(pos(before), '字');
          action();
        }),
      ).toThrow();
      expect(editor.body).toBe(before);
      expect(editor.canUndo).toBe(false);
    }
    insert(editor, pos(before), '好');
    expect(editor.canUndo).toBe(true);
  });

  it('事务句柄不能在回调结束后继续使用，异步回调被拒绝', () => {
    const editor = createTextEditor(parse());
    let saved: TextTransaction | undefined;
    editor.tx((t) => {
      saved = t;
    });
    expect(() => saved?.insertText(pos(editor.body), '晚')).toThrow('已结束');
    // JS 调用者可能绕开 undefined 返回类型，运行时仍需防止提前提交。
    const asyncCallback = (async (t: TextTransaction) => {
      t.insertText(pos(editor.body), '错');
    }) as unknown as (t: TextTransaction) => undefined;
    expect(() => editor.tx(asyncCallback)).toThrow('同步');
    expect(paragraphText(para(editor.body))).toBe('中文正文');
    expect(editor.canUndo).toBe(false);
  });

  it('空插入与折叠删除不产生历史，不清空 redo', () => {
    const editor = createTextEditor(parse());
    const at = pos(editor.body);
    insert(editor, at, '前');
    editor.undo();
    expect(
      editor.tx((t) => {
        t.insertText(at, '');
        t.deleteRange({ start: at, end: at });
      }),
    ).toBeUndefined();
    expect(editor.canUndo).toBe(false);
    expect(editor.canRedo).toBe(true);
  });

  it('非法位置与反向范围被拒绝', () => {
    const editor = createTextEditor(
      parse('<w:p><w:r><w:t>甲乙</w:t></w:r></w:p><w:p><w:r><w:t>丙丁</w:t></w:r></w:p>'),
    );
    for (const at of [
      pos(editor.body, -1),
      pos(editor.body, 0.5),
      pos(editor.body, 3),
      pos(editor.body, 0, 0, 1),
      { ...pos(editor.body), contentIndex: NaN },
      { nodeId: 'missing', contentIndex: 0, offset: 0 },
    ]) {
      expect(() => insert(editor, at, '错')).toThrow();
    }
    expect(() =>
      editor.tx((t) => {
        t.deleteRange({ start: pos(editor.body, 2), end: pos(editor.body, 1) });
      }),
    ).toThrow('起点');
    expect(editor.canUndo).toBe(false);
  });

  it('删除不能越过对象：图片删了回不来，整次拒绝', () => {
    const editor = createTextEditor(parse('<w:p><w:r><w:t>甲</w:t><w:drawing/><w:t>乙</w:t></w:r></w:p>'));
    const before = editor.body;
    expect(() =>
      editor.tx((t) => {
        t.deleteRange({ start: pos(before), end: pos(before, 1, 0, 2) });
      }),
    ).toThrow();
    expect(editor.body).toBe(before);
  });

  it.each(['<w:tab/>', '<w:br/>'])(
    '删除越过 %s 时换成空文字占住槽位，后面的下标不变，撤销还原',
    (element) => {
      const editor = createTextEditor(parse(`<w:p><w:r><w:t>甲</w:t>${element}<w:t>乙丙</w:t></w:r></w:p>`));
      const before = editor.body;
      const change = must(
        editor.tx((t) => {
          t.deleteRange({ start: pos(before, 1), end: pos(before, 1, 0, 2) });
        }),
      );
      expect(para(editor.body).runs[0]?.content).toEqual([
        { kind: 'text', text: '甲' },
        { kind: 'text', text: '' },
        { kind: 'text', text: '丙' },
      ]);
      // 片段后面那个位置落回空文字的开头，第三片的下标不动
      expect(mapTextPosition(pos(before, 1, 0, 1), change)).toEqual(pos(before, 0, 0, 1));
      expect(mapTextPosition(pos(before, 2, 0, 2), change)).toEqual(pos(before, 1, 0, 2));
      editor.undo();
      expect(editor.body).toBe(before);
      // 只选中片段本身之外（片段后到文字）不动它
      editor.tx((t) => {
        t.deleteRange({ start: pos(before, 1, 0, 1), end: pos(before, 1, 0, 2) });
      });
      expect(para(editor.body).runs[0]?.content[1]).toEqual(para(before).runs[0]?.content[1]);
    },
  );

  it('插入制表位 / 软换行：文字中间切开、边界上不造空文字，后续片段下标与位置随之后移', () => {
    const editor = createTextEditor(parse('<w:p><w:r><w:t>甲乙</w:t><w:tab/><w:t>丙</w:t></w:r></w:p>'));
    const before = editor.body;
    let after: DocPosition | undefined;
    const change = must(
      editor.tx((t) => {
        after = t.insertInline(pos(before, 1), 'tab');
      }),
    );
    expect(para(editor.body).runs[0]?.content.map((c) => c.kind)).toEqual([
      'text',
      'tab',
      'text',
      'tab',
      'text',
    ]);
    expect(after).toEqual(pos(before, 1, 0, 1));
    expect(mapTextPosition(pos(before, 2), change)).toEqual(pos(before, 1, 0, 2));
    expect(mapTextPosition(pos(before, 1, 0, 2), change)).toEqual(pos(before, 1, 0, 4));
    expect(mapTextPosition(pos(before, 0, 0, 1), change)).toEqual(pos(before, 0, 0, 3));
    // 撤销：新片段上的位置回到插入点
    const undo = must(editor.undo());
    expect(mapTextPosition(must(after), undo)).toEqual(pos(before, 1));
    expect(editor.body).toBe(before);
    // 文字末尾插软换行，紧接着打字：落进后面已有的文字，不另造片段
    editor.tx((t) => {
      const at = t.insertInline(pos(before, 2), 'lineBreak');
      t.insertText(at, '丁');
    });
    expect(para(editor.body).runs[0]?.content).toEqual([
      { kind: 'text', text: '甲乙' },
      { kind: 'break', breakType: 'line' },
      { kind: 'text', text: '丁' },
      { kind: 'tab' },
      { kind: 'text', text: '丙' },
    ]);
  });

  it('空段落插制表位后接着打字、拆段；制表位后的位置能拆段', () => {
    const editor = createTextEditor(parse('<w:p/>'));
    const paragraph = para(editor.body);
    editor.tx((t) => {
      let at = t.insertInline({ nodeId: paragraph.id, contentIndex: 0, offset: 0 }, 'tab');
      at = t.insertText(at, '甲');
      at = t.insertInline(at, 'tab');
      t.splitParagraph(at);
    });
    expect([...walkParagraphs(editor.body)].map((p) => p.runs.flatMap((r) => r.content))).toEqual([
      [{ kind: 'tab' }, { kind: 'text', text: '甲' }, { kind: 'tab' }, { kind: 'text', text: '' }],
      [{ kind: 'text', text: '' }],
    ]);
    expect(() =>
      editor.tx((t) => {
        t.insertInline(pos(editor.body), 'page' as never);
      }),
    ).toThrow('未知');
  });

  it.each([
    '<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>',
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>',
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p><w:p><w:r><w:t>1</w:t></w:r>',
  ])('保护简单域、复杂域和跨段未闭合域', (xml) => {
    const editor = createTextEditor(parse(`<w:p>${xml}</w:p>`));
    const run = must(
      [...walkParagraphs(editor.body)]
        .flatMap((p) => p.runs)
        .find((r) => r.content.some((c) => c.kind === 'text')),
    );
    expect(() => insert(editor, { nodeId: run.id, contentIndex: 0, offset: 0 }, '改')).toThrow('域');
  });

  it('UTF-16 位置不切开增补字符，输入不混入结构控制字符或孤立代理项', () => {
    const editor = createTextEditor(parse('<w:p><w:r><w:t>甲😀乙</w:t></w:r></w:p>'));
    expect(() => insert(editor, pos(editor.body, 2), '错')).toThrow('代理对');
    for (const text of ['\n', '\r', '\t', '\ud800'])
      expect(() => insert(editor, pos(editor.body), text)).toThrow();
    editor.tx((t) => {
      t.deleteRange({ start: pos(editor.body, 1), end: pos(editor.body, 3) });
    });
    expect(paragraphText(para(editor.body))).toBe('甲乙');
    editor.undo();
    expect(paragraphText(para(editor.body))).toBe('甲😀乙');
  });
});

describe('撤销合并与位置映射', () => {
  it('相邻输入在时间窗内合并，提交返回的变更仍只含本次输入', () => {
    let time = 0;
    const editor = createTextEditor(parse(), { now: () => time });
    const at = pos(editor.body);
    insert(editor, at, '中', true);
    time = 500;
    const second = insert(editor, { ...at, offset: 1 }, '文', true);
    expect(second.changes).toHaveLength(1);
    expect(editor.undo()?.changes.map((c) => c.deletedText)).toEqual(['文', '中']);
    expect(editor.canUndo).toBe(false);
    editor.redo();
    expect(paragraphText(para(editor.body))).toBe('中文中文正文');
  });

  it.each(['timeout', 'backward-clock', 'position', 'command', 'break', 'undo'] as const)(
    '%s 中断连续输入合并',
    (reason) => {
      let time = 1000;
      const editor = createTextEditor(parse(), { now: () => time });
      const at = pos(editor.body);
      insert(editor, at, '甲', true);
      if (reason === 'timeout') time += 1001;
      if (reason === 'backward-clock') time--;
      if (reason === 'break') editor.breakHistory();
      if (reason === 'undo') {
        editor.undo();
        editor.redo();
      }
      insert(editor, { ...at, offset: reason === 'position' ? 2 : 1 }, '乙', reason !== 'command');
      editor.undo();
      expect(paragraphText(para(editor.body))).toBe('甲中文正文');
      expect(editor.canUndo).toBe(true);
    },
  );

  it('撤销后编辑清空 redo；历史容量限制与关闭历史均有效', () => {
    const editor = createTextEditor(parse(), { historyLimit: 1 });
    const at = pos(editor.body);
    insert(editor, at, '甲');
    insert(editor, at, '乙');
    editor.undo();
    expect(paragraphText(para(editor.body))).toBe('甲中文正文');
    expect(editor.canUndo).toBe(false);
    insert(editor, at, '丙');
    expect(editor.canRedo).toBe(false);
    const disabled = createTextEditor(parse(), { historyLimit: 0 });
    insert(disabled, pos(disabled.body), '新', true);
    expect(disabled.canUndo).toBe(false);
    expect(disabled.undo()).toBeUndefined();
  });

  it('位置按本次修改依次平移，撤销反向平移，插入点可选择左右归属', () => {
    const editor = createTextEditor(parse());
    const at = pos(editor.body, 1);
    const changes = insert(editor, at, '新增');
    expect(mapTextPosition(at, changes, 'before')).toEqual(at);
    expect(mapTextPosition(at, changes)).toEqual({ ...at, offset: 3 });
    const moved = mapTextPosition(pos(editor.body, 3), changes);
    expect(moved.offset).toBe(5);
    expect(mapTextPosition(moved, must(editor.undo()))).toEqual(pos(editor.body, 3));
    expect(structuredClone(changes)).toEqual(changes);
  });

  it('选区排除边界新输入，范围被替换后折叠且不翻转', () => {
    const editor = createTextEditor(parse());
    const range: DocRange = { start: pos(editor.body, 1), end: pos(editor.body, 3) };
    const changes = insert(editor, range.start, '前');
    expect(mapTextRange(range, changes)).toEqual({ start: pos(editor.body, 2), end: pos(editor.body, 4) });
    editor.undo();
    const replacement = must(
      editor.tx((t) => {
        const at = t.deleteRange(range);
        t.insertText(at, '替换');
      }),
    );
    expect(mapTextRange(range, replacement)).toEqual({
      start: pos(editor.body, 3),
      end: pos(editor.body, 3),
    });
    expect(mapTextRange({ start: range.start, end: range.start }, changes)).toEqual({
      start: pos(editor.body, 2),
      end: pos(editor.body, 2),
    });
  });

  it('非法历史配置和重复 run id 在初始化时拒绝', () => {
    for (const historyLimit of [-1, 0.5, Infinity])
      expect(() => createTextEditor(parse(), { historyLimit })).toThrow();
    for (const mergeDelay of [-1, NaN, Infinity])
      expect(() => createTextEditor(parse(), { mergeDelay })).toThrow();
    const body = parse();
    para(body).runs.push(structuredClone(must(para(body).runs[0])));
    expect(() => createTextEditor(body)).toThrow('重复');
  });

  it('真实公文编辑后能重新级联与查找，撤销后与原模型完全一致', () => {
    const bytes = new Uint8Array(
      readFileSync(new URL('../../../apps/fidelity/fixtures/gongwen-01.docx', import.meta.url)),
    );
    const loaded = loadDocument(OpcPackage.open(bytes), createDiagnosticSink());
    const editor = createTextEditor(loaded.body);
    const paragraph = must(
      [...walkParagraphs(editor.body)].find((p) =>
        p.runs.some((r) => r.content.some((c) => c.kind === 'text')),
      ),
    );
    const run = must(paragraph.runs.find((r) => r.content.some((c) => c.kind === 'text')));
    const at = { nodeId: run.id, contentIndex: run.content.findIndex((c) => c.kind === 'text'), offset: 0 };
    insert(editor, at, '事务验证');
    const resolved = resolveBody(loaded.cascade, editor.body);
    expect(findText(resolved, '事务验证')).toEqual([{ start: at, end: { ...at, offset: 4 } }]);
    editor.undo();
    expect(editor.body).toEqual(loaded.body);
    expect(resolveBody(loaded.cascade, editor.body)).toEqual(resolveBody(loaded.cascade, loaded.body));
  });
});

describe('段落结构事务', () => {
  it('空段落可输入、拆段并撤销回无 run 原树', () => {
    const editor = createTextEditor(parse('<w:p/>'));
    const before = editor.body;
    const at = { nodeId: para(before).id, contentIndex: 0, offset: 0 };
    let caret = at;
    const change = editor.tx((t) => {
      caret = t.insertText(at, '甲');
      caret = t.splitParagraph(caret);
      caret = t.insertText(caret, '乙');
    });
    expect([...walkParagraphs(editor.body)].map(paragraphText)).toEqual(['甲', '乙']);
    expect(mapTextPosition(at, must(change))).toEqual(caret);
    expect(mapTextPosition(caret, must(editor.undo()))).toEqual(at);
    expect(editor.body).toBe(before);
    editor.redo();
    expect([...walkParagraphs(editor.body)].map(paragraphText)).toEqual(['甲', '乙']);
  });
  it('拆段迁移后续片段，保留样式与链接，映射可逆', () => {
    const editor = createTextEditor(
      parse(
        '<w:p><w:hyperlink r:id="rId1"><w:r><w:rPr><w:b/></w:rPr><w:t>甲乙</w:t><w:t>丙丁</w:t></w:r></w:hyperlink><w:r><w:t>戊</w:t></w:r></w:p>',
      ),
    );
    const before = editor.body;
    const at = pos(before, 1);
    const tail = pos(before, 1, 0, 1);
    const change = must(
      editor.tx((t) => {
        t.splitParagraph(at);
      }),
    );
    expect([...walkParagraphs(editor.body)].map(paragraphText)).toEqual(['甲', '乙丙丁戊']);
    expect(para(editor.body, 1).runs[0]?.hyperlink).toEqual(para(before).runs[0]?.hyperlink);
    expect(para(editor.body, 1).runs[0]?.props).toEqual(para(before).runs[0]?.props);
    const mapped = mapTextPosition(tail, change);
    expect(mapped.contentIndex).toBe(1);
    expect(mapTextPosition(mapped, must(editor.undo()))).toEqual(tail);
    expect(editor.body).toBe(before);
  });
  it('跨三段删除合为一段，保留前段格式，撤销恢复每段', () => {
    const editor = createTextEditor(
      parse('<w:p><w:r><w:t>甲乙</w:t></w:r></w:p><w:p/><w:p><w:r><w:t>丙丁</w:t></w:r></w:p>'),
    );
    const before = editor.body;
    const tail = pos(before, 2, 0, 0, 2);
    const change = must(
      editor.tx((t) => {
        t.deleteRange({ start: pos(before, 1), end: pos(before, 1, 0, 0, 2) });
      }),
    );
    expect([...walkParagraphs(editor.body)].map(paragraphText)).toEqual(['甲丁']);
    expect(mapTextPosition(tail, change).offset).toBe(1);
    editor.undo();
    expect(editor.body).toBe(before);
  });
  it('空段与有字段合并返回有效插入点，两个空段合并保留前段锚点', () => {
    for (const xml of ['<w:p/><w:p><w:r><w:t>甲</w:t></w:r></w:p>', '<w:p/><w:p/>']) {
      const editor = createTextEditor(parse(xml));
      editor.tx((t) => {
        const at = t.joinParagraph(para(editor.body).id);
        t.insertText(at, '前');
      });
      expect(paragraphText(para(editor.body)).startsWith('前')).toBe(true);
    }
  });
  it('跨表格结构失败回滚此前插入，禁止切开代理对与域', () => {
    const editor = createTextEditor(
      parse(
        '<w:p><w:r><w:t>甲</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>乙</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
      ),
    );
    const before = editor.body;
    expect(() =>
      editor.tx((t) => {
        t.insertText(pos(before), '前');
        t.deleteRange({ start: pos(before), end: pos(before, 1, 0, 0, 1) });
      }),
    ).toThrow('跨');
    expect(editor.body).toBe(before);
    const emoji = createTextEditor(parse('<w:p><w:r><w:t>😀</w:t></w:r></w:p>'));
    expect(() =>
      emoji.tx((t) => {
        t.splitParagraph(pos(emoji.body, 1));
      }),
    ).toThrow('代理对');
  });
});

it('空段落无操作不创建 run 或撤销单元；拆段边界尊重 before 亲和性', () => {
  const editor = createTextEditor(parse('<w:p/>'));
  const at = { nodeId: para(editor.body).id, contentIndex: 0, offset: 0 };
  const before = editor.body;
  expect(
    editor.tx((t) => {
      t.insertText(at, '');
      t.deleteRange({ start: at, end: at });
    }),
  ).toBeUndefined();
  expect(editor.body).toBe(before);
  let run = at;
  editor.tx((t) => {
    run = t.insertText(at, '甲');
  });
  const start = { ...run, offset: 0 };
  const change = must(
    editor.tx((t) => {
      t.splitParagraph(start);
    }),
  );
  expect(mapTextPosition(start, change, 'before')).toEqual(start);
  expect(mapTextPosition(start, change, 'after').nodeId).not.toBe(start.nodeId);
});

it('预提交检查失败保留模型及 undo/redo 历史', () => {
  let reject = false;
  const editor = createTextEditor(parse(), {
    validate: () => {
      if (reject) throw new Error('重排失败');
    },
  });
  insert(editor, pos(editor.body), '前');
  const before = editor.body;
  reject = true;
  expect(() => insert(editor, pos(editor.body), '错')).toThrow('重排失败');
  expect(() => editor.undo()).toThrow('重排失败');
  expect(editor.body).toBe(before);
  expect(editor.canUndo).toBe(true);
  reject = false;
  editor.undo();
  expect(editor.canRedo).toBe(true);
});
