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

describe('插列 / 删列', () => {
  const grid2 = '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>';
  const tblW = (w: number) => `<w:tblPr><w:tblW w:w="${w}" w:type="dxa"/></w:tblPr>`;
  const w = (n: number) => `<w:tcW w:w="${n}" w:type="dxa"/>`;
  const table2 = (...rows: string[]) => `<w:tbl>${tblW(5000)}${grid2}${rows.join('')}</w:tbl>`;
  const widths = (t: Table) => t.rows.map((r) => r.cells.map((c) => [c.gridSpan, c.props.width?.value]));

  it('右插照左邻抄、网格与 dxa 表宽跟着长，光标进本行新格；撤销整表回去', () => {
    const body = parse(
      table2(
        tr(tc(p('甲', '<w:jc w:val="center"/>'), w(2000)), tc(p('乙'), w(3000))),
        tr(tc(p('丙'), w(2000)), tc(p('丁'), w(3000))),
      ),
    );
    const editor = createTextEditor(body);
    let caret: DocPosition | undefined;
    editor.tx((t) => {
      caret = t.insertColumn(at(body, '丙'), 'right');
    });
    const table = tables(editor.body)[0] as Table;
    expect(grid(table)).toEqual([
      ['甲', '·', '乙'],
      ['丙', '·', '丁'],
    ]);
    expect(table.grid).toEqual([2000, 2000, 3000]);
    expect(table.props.width).toEqual({ value: 7000, type: 'dxa' });
    expect(widths(table)[0]).toEqual([
      [1, 2000],
      [1, 2000],
      [1, 3000],
    ]);
    expect(table.rows[0]?.cells[1]?.blocks[0]).toMatchObject({ props: { justification: 'center' } });
    expect(caret).toEqual({ nodeId: table.rows[1]?.cells[1]?.blocks[0]?.id, contentIndex: 0, offset: 0 });
    const back = editor.undo();
    expect(mapTextPosition(caret as DocPosition, back as never)).toEqual(at(body, '丙'));
    expect(editor.body).toEqual(createTextEditor(body).body);
  });

  it('左插宽度取本列；跨过边界的格被撑宽，挨着纵向合并区的新格照样合并', () => {
    const body = parse(
      table2(
        tr(tc(p('横'), `${w(5000)}<w:gridSpan w:val="2"/>`)),
        tr(tc(p('纵'), `${w(2000)}<w:vMerge w:val="restart"/>`), tc(p('乙'), w(3000))),
        tr(tc(p(''), `${w(2000)}<w:vMerge/>`), tc(p('丁'), w(3000))),
      ),
    );
    const editor = createTextEditor(body);
    editor.tx((t) => {
      t.insertColumn(at(body, '乙'), 'left');
    });
    const table = tables(editor.body)[0] as Table;
    expect(table.grid).toEqual([2000, 3000, 3000]);
    expect(grid(table)).toEqual([['横'], ['纵', '·', '乙'], ['^', '·', '丁']]);
    expect(widths(table)[0]).toEqual([[3, 8000]]);
    // 左插抄的是右边那格（乙），它不在合并区里
    expect(table.rows.slice(1).map((r) => r.cells[1]?.vMerge)).toEqual(['none', 'none']);
    const right = createTextEditor(body);
    right.tx((t) => {
      t.insertColumn(at(body, '纵'), 'right');
    });
    expect(
      tables(right.body)[0]
        ?.rows.slice(1)
        .map((r) => r.cells[1]?.vMerge),
    ).toEqual(['restart', 'continue']);
  });

  it('删列：整格删、跨列格缩窄，被删格里的光标收拢到接替它的格', () => {
    const body = parse(
      `<w:tbl>${tblW(7000)}<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
        tr(tc(p('横'), `${w(4000)}<w:gridSpan w:val="2"/>`), tc(p('尾'), w(3000))) +
        tr(tc(p('甲'), w(2000)), tc(p('乙'), w(2000)), tc(p('丙'), w(3000))) +
        '</w:tbl>',
    );
    const editor = createTextEditor(body);
    let caret: DocPosition | undefined;
    const changes = editor.tx((t) => {
      caret = t.deleteColumns({ start: at(body, '乙'), end: at(body, '乙') });
    });
    const table = tables(editor.body)[0] as Table;
    expect(table.grid).toEqual([2000, 3000]);
    expect(table.props.width).toEqual({ value: 5000, type: 'dxa' });
    expect(grid(table)).toEqual([
      ['横', '尾'],
      ['甲', '丙'],
    ]);
    expect(widths(table)[0]).toEqual([
      [1, 2000],
      [1, 3000],
    ]);
    expect(caret).toEqual(at(body, '丙'));
    expect(mapTextPosition(at(body, '乙', 1), changes as never)).toEqual(at(body, '丙'));
    editor.undo();
    expect(editor.body).toEqual(createTextEditor(body).body);
  });

  it('删掉合并区首格那一列时下面的续格升成 restart；删空的行删掉，删光删表', () => {
    const body = parse(
      table2(
        tr(tc(p('甲'), w(2000)), tc(p('乙'), w(3000))),
        tr(tc(p('单'), `${w(5000)}<w:gridSpan w:val="2"/>`)),
      ) + p('表后'),
    );
    const editor = createTextEditor(body);
    editor.tx((t) => {
      t.deleteColumns({ start: at(body, '甲'), end: at(body, '甲') });
    });
    expect(grid(tables(editor.body)[0] as Table)).toEqual([['乙'], ['单']]);
    let caret: DocPosition | undefined;
    editor.tx((t) => {
      caret = t.deleteColumns({ start: at(body, '乙'), end: at(body, '乙') });
    });
    expect(tables(editor.body)).toHaveLength(0);
    expect(caret).toEqual(at(body, '表后'));

    // 首格只占第一列、续格跨两列：删第一列后首格没了，续格缩成一列，上面同列是个普通格
    const merged = parse(
      table2(
        tr(tc(p('首'), `${w(2000)}<w:vMerge w:val="restart"/>`), tc(p('乙'), w(3000))),
        tr(tc(p(''), `${w(5000)}<w:gridSpan w:val="2"/><w:vMerge/>`)),
      ),
    );
    const m = createTextEditor(merged);
    m.tx((t) => {
      t.deleteColumns({ start: at(merged, '首'), end: at(merged, '首') });
    });
    const repaired = tables(m.body)[0] as Table;
    expect(repaired.rows.map((r) => r.cells.map((c) => [c.vMerge, c.gridSpan]))).toEqual([
      [['none', 1]],
      [['restart', 1]],
    ]);
  });

  it('没有网格的表拒绝插列；两端不在同一张表拒绝删列', () => {
    const body = parse(tbl(tr(tc(p('甲')))) + tbl(tr(tc(p('乙')))));
    const editor = createTextEditor(body);
    expect(() =>
      editor.tx((t) => {
        t.insertColumn(at(body, '甲'), 'left');
      }),
    ).toThrow(/tblGrid/);
    expect(() =>
      editor.tx((t) => {
        t.deleteColumns({ start: at(body, '甲'), end: at(body, '乙') });
      }),
    ).toThrow(/同一张表/);
  });
});

describe('合并 / 拆分单元格', () => {
  const w = (n: number) => `<w:tcW w:w="${n}" w:type="dxa"/>`;
  const grid3 =
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>';
  const t3 = (...rows: string[]) => `<w:tbl>${grid3}${rows.join('')}</w:tbl>`;
  const texts = (cell: { blocks: unknown[] } | undefined) =>
    (cell?.blocks ?? []).map((b) => paragraphText(b as never));
  const shape = (t: Table) =>
    t.rows.map((r) => r.cells.map((c) => [c.gridSpan, c.vMerge, c.props.width?.value]));

  it('横向合并：内容按左到右接进首格、空格不贡献段落，宽按网格重算；段落 id 不变，位置不用映射', () => {
    const body = parse(t3(tr(tc(p('甲'), w(1000)), tc(p(''), w(2000)), tc(p('丙') + p('丙二'), w(3000)))));
    const editor = createTextEditor(body);
    let caret: DocPosition | undefined;
    const changes = editor.tx((t) => {
      caret = t.mergeCells({ start: at(body, '甲'), end: at(body, '丙') });
    });
    const table = tables(editor.body)[0] as Table;
    expect(shape(table)).toEqual([[[3, 'none', 6000]]]);
    expect(texts(table.rows[0]?.cells[0])).toEqual(['甲', '丙', '丙二']);
    expect(caret).toEqual(at(body, '甲'));
    expect(mapTextPosition(at(body, '丙', 1), changes as never)).toEqual(at(body, '丙', 1));
    editor.undo();
    expect(editor.body).toEqual(createTextEditor(body).body);
  });

  it('跨两行合并：首格 restart、下面留跨同样列数的续格；内容按行接起来', () => {
    const body = parse(
      t3(
        tr(tc(p('甲'), w(1000)), tc(p('乙'), w(2000)), tc(p('外一'), w(3000))),
        tr(tc(p('丙'), w(1000)), tc(p(''), w(2000)), tc(p('外二'), w(3000))),
      ),
    );
    const editor = createTextEditor(body);
    editor.tx((t) => {
      t.mergeCells({ start: at(body, '甲'), end: at(body, '丙') });
    });
    const table = tables(editor.body)[0] as Table;
    // 只选了第一列 —— 矩形就是第一列两行
    expect(shape(table)).toEqual([
      [
        [1, 'restart', 1000],
        [1, 'none', 2000],
        [1, 'none', 3000],
      ],
      [
        [1, 'continue', 1000],
        [1, 'none', 2000],
        [1, 'none', 3000],
      ],
    ]);
    expect(texts(table.rows[0]?.cells[0])).toEqual(['甲', '丙']);
    expect(texts(table.rows[1]?.cells[0])).toEqual(['']);

    const wide = createTextEditor(body);
    wide.tx((t) => {
      t.mergeCells({ start: at(body, '甲'), end: { ...at(body, '外二'), offset: 1 } });
    });
    const all = tables(wide.body)[0] as Table;
    expect(shape(all)).toEqual([[[3, 'restart', 6000]], [[3, 'continue', 6000]]]);
    expect(texts(all.rows[0]?.cells[0])).toEqual(['甲', '乙', '外一', '丙', '外二']);
    wide.undo();
    expect(wide.body).toEqual(createTextEditor(body).body);
  });

  it('选区被合并格撑大：碰到跨列格与纵向合并区就把它们整个包进来', () => {
    const body = parse(
      t3(
        tr(
          tc(p('横'), `${w(3000)}<w:gridSpan w:val="2"/>`),
          tc(p('纵'), `${w(3000)}<w:vMerge w:val="restart"/>`),
        ),
        tr(tc(p('丙'), w(1000)), tc(p('丁'), w(2000)), tc(p(''), `${w(3000)}<w:vMerge/>`)),
      ),
    );
    const editor = createTextEditor(body);
    // 只点了「丁」与「纵」：纵向合并区往下撑到第二行、跨列格往左撑到第一列
    editor.tx((t) => {
      t.mergeCells({ start: at(body, '纵'), end: at(body, '丁') });
    });
    const table = tables(editor.body)[0] as Table;
    expect(shape(table)).toEqual([[[3, 'restart', 6000]], [[3, 'continue', 6000]]]);
    expect(texts(table.rows[0]?.cells[0])).toEqual(['横', '纵', '丙', '丁']);
  });

  it('跳过的网格列撑不成矩形时拒绝；只有一格时不修改', () => {
    const body = parse(
      t3(
        tr(tc(p('甲'), w(1000)), tc(p('乙'), w(2000)), tc(p('丙'), w(3000))),
        `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>${tc(p('丁'), w(2000))}${tc(p('戊'), w(3000))}</w:tr>`,
      ),
    );
    const editor = createTextEditor(body);
    expect(() =>
      editor.tx((t) => {
        t.mergeCells({ start: at(body, '甲'), end: at(body, '丁') });
      }),
    ).toThrow(/矩形/);
    expect(
      editor.tx((t) => {
        t.mergeCells({ start: at(body, '甲'), end: { ...at(body, '甲'), offset: 1 } });
      }),
    ).toBeUndefined();
  });

  it('拆分是合并的逆操作：合并再拆回，结构与原表一致，内容留在首格', () => {
    const body = parse(
      t3(
        tr(tc(p('甲'), w(1000)), tc(p('乙'), w(2000)), tc(p('丙'), w(3000))),
        tr(tc(p(''), w(1000)), tc(p(''), w(2000)), tc(p('丁'), w(3000))),
      ),
    );
    const editor = createTextEditor(body);
    editor.tx((t) => {
      t.mergeCells({ start: at(body, '甲'), end: { ...at(body, '乙'), offset: 1 } });
      t.mergeCells({ start: at(body, '甲'), end: at(body, '丁') });
    });
    expect(shape(tables(editor.body)[0] as Table)).toEqual([[[3, 'restart', 6000]], [[3, 'continue', 6000]]]);
    let caret: DocPosition | undefined;
    editor.tx((t) => {
      caret = t.splitCell(at(body, '甲'));
    });
    const table = tables(editor.body)[0] as Table;
    expect(shape(table)).toEqual(shape(tables(createTextEditor(body).body)[0] as Table));
    expect(texts(table.rows[0]?.cells[0])).toEqual(['甲', '乙', '丙', '丁']);
    expect(caret).toEqual(at(body, '甲'));
    expect(
      editor.tx((t) => {
        t.splitCell(at(body, '甲'));
      }),
    ).toBeUndefined();
  });
});
