import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { Body, Table } from './nodes.ts';
import { paragraphText, walkBlocks, walkParagraphs } from './nodes.ts';
import { parseBody } from './parse-body.ts';
import type { DocPosition } from './position.ts';
import { mapTextPosition } from './text-change.ts';
import { createTextEditor } from './text-transaction.ts';

function parse(xml: string): Body {
  return parseBody(parseXml(`<w:document><w:body>${xml}</w:body></w:document>`), createDiagnosticSink());
}
const p = (text: string, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p>`;
const tc = (inner: string, pr = '') => `<w:tc>${pr ? `<w:tcPr>${pr}</w:tcPr>` : ''}${inner}</w:tc>`;
const tr = (...cells: string[]) => `<w:tr>${cells.join('')}</w:tr>`;
const tbl = (...rows: string[]) => `<w:tbl>${rows.join('')}</w:tbl>`;

function at(body: Body, text: string, offset = 0): DocPosition {
  for (const para of walkParagraphs(body))
    for (const run of para.runs)
      if (run.content.some((c) => c.kind === 'text' && c.text === text))
        return { nodeId: run.id, contentIndex: 0, offset };
  throw new Error(`没有「${text}」`);
}
function tables(body: Body): Table[] {
  return body.sections.flatMap((s) =>
    [...walkBlocks(s.blocks)].filter((b): b is Table => b.kind === 'table'),
  );
}
/** 每行每格的首段文字，续格记成「^」—— 一眼看出合并区长什么样 */
function grid(table: Table): string[][] {
  return table.rows.map((r) =>
    r.cells.map((c) => (c.vMerge === 'continue' ? '^' : paragraphText(c.blocks[0] as never) || '·')),
  );
}

describe('插行', () => {
  const body = parse(
    p('表前') +
      tbl(
        tr(tc(p('甲', '<w:jc w:val="center"/>'), '<w:tcW w:w="3000" w:type="dxa"/>'), tc(p('乙'))),
        tr(tc(p('丙')), tc(p('丁'))),
      ) +
      p('表后'),
  );

  it('照所在行抄格数与格属性，段落格式抄模板格首段，光标在新行首格', () => {
    const editor = createTextEditor(body);
    let caret: DocPosition | undefined;
    const changes = editor.tx((t) => {
      caret = t.insertRow(at(body, '乙'), 'below');
    });
    const [table] = tables(editor.body);
    expect(grid(table as Table)).toEqual([
      ['甲', '乙'],
      ['·', '·'],
      ['丙', '丁'],
    ]);
    const row = (table as Table).rows[1];
    expect(row?.cells[0]?.props).toEqual((table as Table).rows[0]?.cells[0]?.props);
    expect(row?.cells[0]?.blocks[0]).toMatchObject({ props: { justification: 'center' }, runs: [] });
    expect(caret).toEqual({ nodeId: row?.cells[0]?.blocks[0]?.id, contentIndex: 0, offset: 0 });
    expect(changes?.paragraphIds).toHaveLength(2);
    // 光标就在新格里：接着打字即写进空段落
    editor.tx((t) => {
      t.insertText(caret as DocPosition, '新');
    });
    expect(grid(tables(editor.body)[0] as Table)[1]).toEqual(['新', '·']);
    editor.undo();
    // 撤销插行：新行里的光标收拢回按 Tab 的那个位置，不留在已经不存在的段落上
    const back = editor.undo();
    expect(mapTextPosition(caret as DocPosition, back as never)).toEqual(at(body, '乙'));
    expect(editor.body).toEqual(createTextEditor(body).body);
  });

  it('above 插在所在行前面；不在表格里抛错且整次回滚', () => {
    const editor = createTextEditor(body);
    editor.tx((t) => {
      t.insertRow(at(body, '丙'), 'above');
    });
    expect(grid(tables(editor.body)[0] as Table).map((r) => r[0])).toEqual(['甲', '·', '丙']);
    const before = editor.body;
    expect(() =>
      editor.tx((t) => {
        t.insertRow(at(body, '表前'), 'below');
      }),
    ).toThrow(/不在表格/);
    expect(editor.body).toBe(before);
  });

  it('纵向合并按网格列对齐：插在合并区中间写 continue，插在区外写 none', () => {
    // 第一列三行合并；第二行有一格跨两列，按下标比会把「戊」当成第一列
    const merged = parse(
      tbl(
        tr(tc(p('首'), '<w:vMerge w:val="restart"/>'), tc(p('乙')), tc(p('丙'))),
        tr(tc(p(''), '<w:vMerge/>'), tc(p('戊'), '<w:gridSpan w:val="2"/>')),
        tr(tc(p(''), '<w:vMerge/>'), tc(p('己')), tc(p('庚'))),
      ),
    );
    const editor = createTextEditor(merged);
    editor.tx((t) => {
      t.insertRow(at(merged, '乙'), 'below');
      t.insertRow(at(merged, '己'), 'below');
      t.insertRow(at(merged, '乙'), 'above');
    });
    expect(grid(tables(editor.body)[0] as Table)).toEqual([
      ['·', '·', '·'], // 首格上方：合并区外
      ['首', '乙', '丙'],
      ['^', '·', '·'], // 首格下方、下一行还是续格：合并区长一格
      ['^', '戊'],
      ['^', '己', '庚'],
      ['·', '·', '·'], // 合并区最后一行下方：区外
    ]);
  });

  it('嵌套表格按最内层那张插', () => {
    const nested = parse(tbl(tr(tc(p('外') + tbl(tr(tc(p('内一')), tc(p('内二'))))))));
    const editor = createTextEditor(nested);
    editor.tx((t) => {
      t.insertRow(at(nested, '内二'), 'below');
    });
    const [outer, inner] = tables(editor.body);
    expect(outer?.rows).toHaveLength(1);
    expect(inner?.rows).toHaveLength(2);
  });
});

describe('删行', () => {
  it('删范围两端之间的行，位置收拢到下一行首格；撤销恢复', () => {
    const body = parse(tbl(tr(tc(p('甲')), tc(p('乙'))), tr(tc(p('丙')), tc(p('丁'))), tr(tc(p('戊')))));
    const editor = createTextEditor(body);
    let caret: DocPosition | undefined;
    const changes = editor.tx((t) => {
      caret = t.deleteRows({ start: at(body, '乙', 1), end: at(body, '丙') });
    });
    expect(grid(tables(editor.body)[0] as Table)).toEqual([['戊']]);
    expect(caret).toEqual(at(body, '戊'));
    // 被删行里的位置（第 1 个字之后）收拢到目标，不按偏移平移
    expect(mapTextPosition(at(body, '丁', 1), changes as never)).toEqual(at(body, '戊'));
    editor.undo();
    expect(grid(tables(editor.body)[0] as Table)).toHaveLength(3);
  });

  it('删最后一行光标落到上一行；删光整张表落到表后，容器空了补一个空段落', () => {
    const body = parse(tbl(tr(tc(p('甲'))), tr(tc(p('乙')))) + p('表后'));
    const editor = createTextEditor(body);
    let caret: DocPosition | undefined;
    editor.tx((t) => {
      caret = t.deleteRows({ start: at(body, '乙'), end: at(body, '乙') });
    });
    expect(caret).toEqual(at(body, '甲'));
    editor.tx((t) => {
      caret = t.deleteRows({ start: at(body, '甲'), end: at(body, '甲') });
    });
    expect(tables(editor.body)).toHaveLength(0);
    expect(caret).toEqual(at(body, '表后'));

    const nested = parse(tbl(tr(tc(tbl(tr(tc(p('内'))))))));
    const inner = createTextEditor(nested);
    inner.tx((t) => {
      caret = t.deleteRows({ start: at(nested, '内'), end: at(nested, '内') });
    });
    const cell = tables(inner.body)[0]?.rows[0]?.cells[0];
    expect(cell?.blocks).toHaveLength(1);
    expect(cell?.blocks[0]).toMatchObject({ kind: 'paragraph', runs: [] });
    expect(caret).toEqual({ nodeId: cell?.blocks[0]?.id, contentIndex: 0, offset: 0 });
  });

  it('删掉合并区首格所在行，下面的续格升成 restart；删区中间的行不动', () => {
    const body = parse(
      tbl(
        tr(tc(p('首'), '<w:vMerge w:val="restart"/>'), tc(p('乙'))),
        tr(tc(p(''), '<w:vMerge/>'), tc(p('丁'))),
        tr(tc(p(''), '<w:vMerge/>'), tc(p('己'))),
      ),
    );
    const middle = createTextEditor(body);
    middle.tx((t) => {
      t.deleteRows({ start: at(body, '丁'), end: at(body, '丁') });
    });
    expect(tables(middle.body)[0]?.rows.map((r) => r.cells[0]?.vMerge)).toEqual(['restart', 'continue']);
    const head = createTextEditor(body);
    head.tx((t) => {
      t.deleteRows({ start: at(body, '乙'), end: at(body, '乙') });
    });
    expect(tables(head.body)[0]?.rows.map((r) => r.cells[0]?.vMerge)).toEqual(['restart', 'continue']);
  });

  it('两端不在同一张表、或行里有域时拒绝', () => {
    const body = parse(
      tbl(tr(tc(p('甲')))) +
        tbl(
          tr(
            tc(
              '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r>' +
                '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>' +
                '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
            ),
          ),
          tr(tc(p('乙'))),
        ),
    );
    const editor = createTextEditor(body);
    expect(() =>
      editor.tx((t) => {
        t.deleteRows({ start: at(body, '甲'), end: at(body, '乙') });
      }),
    ).toThrow(/同一张表/);
    expect(() =>
      editor.tx((t) => {
        t.deleteRows({ start: at(body, '1'), end: at(body, '1') });
      }),
    ).toThrow(/域/);
  });
});
