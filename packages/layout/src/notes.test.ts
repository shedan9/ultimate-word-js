/**
 * 脚注：号怎么数（notes.ts）、分页怎么给它让地方（page.ts）。
 *
 * 版心与 page.test.ts 同一套：「一行 10 个字、一页 3 行」，行高 `EA_LINE` = 273。
 * 脚注内容也是一行，于是「本页放几行正文」能心算：带一条一行的脚注就只剩两行。
 * 分隔线那一段只在一个用例里给（它是个空段落，行高走拉丁规则，不是 273），
 * 其余用例都不给 —— 期望值因此不必依赖空段落的行高。
 *
 * **没有真值**：Word 的脚注区从来没有跟 PDF 比过，这里测的是规则自洽（让出的高 = 画出的高、
 * 引用与脚注同页），不是几何精度。
 */
import { createDiagnosticSink } from '@uw/core';
import type { NoteNumbering, ResolvedBlock, ResolvedBody, RunContent, SectionProps } from '@uw/model';
import { DEFAULT_SECTION_PROPS, DEFAULT_SETTINGS } from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { NoteSource } from './notes.ts';
import { noteLabels } from './notes.ts';
import type { DocumentLayout, LayoutDocumentOptions, PlacedParagraph } from './page.ts';
import { layoutDocument } from './page.ts';
import { cell, fakeMeasurer, NO_GRID, para, row, run, runOf, SIZE_5, table } from './test-fixtures.ts';

const EA_LINE = SIZE_5 * 1.3;

function sect(over: Partial<SectionProps> = {}): SectionProps {
  return {
    ...structuredClone(DEFAULT_SECTION_PROPS),
    page: { width: 3300, height: 2019, orientation: 'portrait' },
    margin: { top: 600, right: 600, bottom: 600, left: 600, header: 0, footer: 0, gutter: 0 },
    docGrid: NO_GRID,
    ...over,
  };
}

function body(blocks: ResolvedBlock[], props: SectionProps = sect()): ResolvedBody {
  return { sections: [{ id: 's0', props, blocks }] };
}

const ref = (
  noteId: string,
  extra: Partial<Extract<RunContent, { kind: 'noteReference' }>> = {},
): RunContent => ({
  kind: 'noteReference',
  noteType: 'footnote',
  noteId,
  ...extra,
});
const mark: RunContent = { kind: 'noteMark', noteType: 'footnote' };

/** 一条脚注：开头的号 + 一段字（`lines` 行） */
function note(text = '注', lines = 1): { resolved: ResolvedBlock[] } {
  const blocks: ResolvedBlock[] = [para([runOf([mark]), run(text)], { widowControl: false })];
  for (let i = 1; i < lines; i++) blocks.push(para([run(text)], { widowControl: false }));
  return { resolved: blocks };
}

function source(
  notes: Record<string, { resolved: ResolvedBlock[] }>,
  over: { numbering?: NoteNumbering; separator?: { resolved: ResolvedBlock[] } } = {},
): NoteSource {
  return {
    footnotes: {
      notes,
      numbering: over.numbering ?? {},
      ...(over.separator === undefined ? {} : { separator: over.separator }),
    },
    endnotes: { notes: {}, numbering: {} },
  };
}

function opts(over: Partial<LayoutDocumentOptions> = {}): LayoutDocumentOptions {
  return { measurer: fakeMeasurer(), settings: DEFAULT_SETTINGS, ...over };
}

/** 一行五个字的段落（可选带一个脚注引用） */
const line = (noteId?: string) =>
  para(noteId === undefined ? [run('一二三四五')] : [run('一二三四五'), runOf([ref(noteId)])], {
    widowControl: false,
  });

const shape = (doc: DocumentLayout): number[] =>
  doc.pages.map((p) =>
    p.blocks
      .filter((b): b is PlacedParagraph => b.kind === 'paragraph')
      .reduce((n, b) => n + b.lines.length, 0),
  );

describe('脚注编号（noteLabels）', () => {
  it('按文档序数，隐藏的与自定义标记的不占号，注内容开头的号跟着那一条', () => {
    const a = runOf([ref('7')]);
    const hidden = runOf([ref('8')], { hidden: true });
    const custom = runOf([ref('9', { customMark: true })]);
    const b = runOf([ref('3')]);
    const n3 = note();
    const labels = noteLabels(
      body([para([run('甲'), a]), para([hidden, custom, run('*'), b])]),
      source({ '7': note(), '3': n3 }),
    );
    expect(labels.get(a.id)).toBe('1');
    expect(labels.get(b.id)).toBe('2');
    expect(labels.has(hidden.id)).toBe(false);
    expect(labels.has(custom.id)).toBe(false);
    // `w:id` 是 3，但它是第二条 —— 号看的是先后不是 id
    const markRun = (n3.resolved[0] as { runs: { id: string }[] }).runs[0] as { id: string };
    expect(labels.get(markRun.id)).toBe('2');
  });

  it('节上的格式盖文档级的，eachSect 每节从 numStart 起', () => {
    const a = runOf([ref('1')]);
    const b = runOf([ref('2')]);
    const c = runOf([ref('3')]);
    const doc: ResolvedBody = {
      sections: [
        { id: 's0', props: sect(), blocks: [para([a, b])] },
        {
          id: 's1',
          props: sect({ footnotePr: { numFmt: 'decimalEnclosedCircleChinese', numRestart: 'eachSect' } }),
          blocks: [para([c])],
        },
      ],
    };
    const labels = noteLabels(doc, source({}, { numbering: { numStart: 5 } }));
    expect([labels.get(a.id), labels.get(b.id), labels.get(c.id)]).toEqual(['5', '6', '⑤']);
  });

  it('每页重新编号还没做：按连续编号数并记诊断', () => {
    const sink = createDiagnosticSink();
    noteLabels(body([line('1')]), source({ '1': note() }, { numbering: { numRestart: 'eachPage' } }), sink);
    expect(sink.list().map((d) => d.code)).toEqual(['note-restart-each-page']);
  });
});

describe('脚注占位（分页）', () => {
  it('不传脚注数据时号是空的、页底不留地方 —— 与没有脚注的文档一样排', () => {
    const doc = layoutDocument(body([line('1'), line(), line()]), opts());
    expect(shape(doc)).toEqual([3]);
    expect(doc.pages[0]?.footnotes).toBeUndefined();
  });

  it('引用所在的那一页让出脚注的高度，脚注区底边贴着版心底', () => {
    const doc = layoutDocument(body([line('1'), line(), line()]), opts({ notes: source({ '1': note() }) }));
    // 一页三行，脚注占掉一行 → 本页只剩两行正文
    expect(shape(doc)).toEqual([2, 1]);
    const page = doc.pages[0];
    const area = page?.footnotes;
    expect(area?.notes).toEqual(['footnote:1']);
    expect(area?.height).toBe(EA_LINE);
    const c = page?.geometry.content;
    expect((area?.y ?? 0) + (area?.height ?? 0)).toBe((c?.y ?? 0) + (c?.height ?? 0));
    expect(area?.x).toBe(c?.x);
    // 脚注内容开头的号与正文里的号一致
    const first = (area?.blocks[0] as PlacedParagraph | undefined)?.lines[0]?.line.fragments[0];
    expect(first?.text).toBe('1');
    expect(doc.pages[1]?.footnotes).toBeUndefined();
  });

  it('带引用的行连同脚注放不下时，这一行跟着脚注一起去下一页', () => {
    const doc = layoutDocument(body([line(), line(), line('1')]), opts({ notes: source({ '1': note() }) }));
    expect(shape(doc)).toEqual([2, 1]);
    expect(doc.pages[0]?.footnotes).toBeUndefined();
    expect(doc.pages[1]?.footnotes?.notes).toEqual(['footnote:1']);
  });

  it('同一页引两次同一条只占一份', () => {
    const doc = layoutDocument(
      body([line('1'), line('1'), line()]),
      opts({ notes: source({ '1': note() }) }),
    );
    // 第二行引的「1」已经在本页，不再加高：两行正文 + 一条脚注 = 三行
    expect(shape(doc)).toEqual([2, 1]);
    expect(doc.pages[0]?.footnotes?.notes).toEqual(['footnote:1']);
  });

  it('两条脚注按引用的先后摞；引到的脚注已在别页的，本页照样要再收一份', () => {
    // 一页五行
    const tall = sect({ page: { width: 3300, height: 1200 + 5 * EA_LINE, orientation: 'portrait' } });
    const doc = layoutDocument(
      body([line('2'), line('1'), line('2'), line('1')], tall),
      opts({ notes: source({ '1': note('甲'), '2': note('乙') }) }),
    );
    // 第一页：三行正文 + 两条脚注 = 五行；第四行引的「1」在第一页收过，到了第二页还得再收
    expect(shape(doc)).toEqual([3, 1]);
    const area = doc.pages[0]?.footnotes;
    expect(area?.notes).toEqual(['footnote:2', 'footnote:1']);
    expect(area?.blocks.map((b) => b.y)).toEqual([0, EA_LINE]);
    expect(doc.pages[1]?.footnotes?.notes).toEqual(['footnote:1']);
  });

  it('分隔线那一段的高度算进让出的地方，线画在区域顶上那一段里', () => {
    const separator = { resolved: [para([run('')], { widowControl: false })] as ResolvedBlock[] };
    const doc = layoutDocument(
      body([line('1'), line()]),
      opts({ notes: source({ '1': note() }, { separator }) }),
    );
    const area = doc.pages[0]?.footnotes;
    expect(area).toBeDefined();
    const sepHeight = (area?.height ?? 0) - EA_LINE;
    expect(sepHeight).toBeGreaterThan(0);
    // 脚注正文从分隔线那一段下面开始，分隔线本身不进 blocks
    expect(area?.blocks).toHaveLength(1);
    expect(area?.blocks[0]?.y).toBe(sepHeight);
    expect(area?.separator.y).toBeCloseTo(sepHeight / 2);
    expect(area?.separator.width).toBe(2100);
  });

  it('表格行里引的脚注同样让地方', () => {
    const make = () =>
      table([2100], [row([cell([line('1')])]), row([cell([line()])]), row([cell([line()])])]);
    const rowsOnFirst = (doc: DocumentLayout) =>
      doc.pages[0]?.blocks.reduce((n, b) => n + (b.kind === 'table' ? b.rows.length : 0), 0) ?? 0;
    const plain = layoutDocument(body([make()]), opts());
    const noted = layoutDocument(body([make()]), opts({ notes: source({ '1': note() }) }));
    expect(rowsOnFirst(noted)).toBeLessThan(rowsOnFirst(plain));
    expect(noted.pages[0]?.footnotes?.notes).toEqual(['footnote:1']);
  });

  it('脚注比一整页还长时整条硬塞并记诊断（跨页续排还没做）', () => {
    const sink = createDiagnosticSink();
    const doc = layoutDocument(
      body([line('1')]),
      opts({ notes: source({ '1': note('长', 4) }), diagnostics: sink }),
    );
    expect(doc.pages[0]?.footnotes?.notes).toEqual(['footnote:1']);
    expect(sink.list().map((d) => d.code)).toContain('footnote-overflow');
  });

  it('正文里的号自成片段：带域结果标记、位置 -1，不与同一 run 里的字并在一起', () => {
    const r = runOf([{ kind: 'text', text: '甲' }, ref('1')]);
    const doc = layoutDocument(body([para([r])]), opts({ notes: source({ '1': note() }) }));
    const placed = doc.pages[0]?.blocks[0] as PlacedParagraph | undefined;
    const frags = placed?.lines[0]?.line.fragments ?? [];
    expect(frags.map((f) => [f.text, f.contentIndex, f.offset, f.field === true])).toEqual([
      ['甲', 0, 0, false],
      ['1', 1, -1, true],
    ]);
    expect(placed?.lines[0]?.line.notes).toEqual(['footnote:1']);
  });
});
