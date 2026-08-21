/**
 * 分页。
 *
 * 版心刻意造成「一行 10 个字、一页 3 行」—— 期望值可以数着字与行写出来，
 * 失败时一眼看得出是差了一个字还是差了一行。行高固定 `EA_LINE`（合成度量器的
 * 东亚行高 = 1.3 em），页内 y 因此是 0 / 273 / 546 这种能心算的数。
 *
 * 真实字体、真实公文的分页精度由 `fixture.test.ts` 的 L3 一节兜着（18 行基线
 * 与 Word 真值最大差 0.06pt）；这里只测**规则**：孤行寡行、keepNext / keepLines、
 * 硬分页符、分节、表格按行拆页。
 */
import { createDiagnosticSink } from '@uw/core';
import type { ResolvedBlock, ResolvedBody, SectionProps } from '@uw/model';
import { DEFAULT_SECTION_PROPS, DEFAULT_SETTINGS } from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { DocumentLayout, LayoutDocumentOptions, PlacedParagraph, PlacedTable } from './page.ts';
import { layoutDocument, pageGeometry } from './page.ts';
import {
  cell,
  fakeMeasurer,
  NO_GRID,
  para,
  paraProps,
  row,
  rowProps,
  run,
  runOf,
  SIZE_5,
  table,
} from './test-fixtures.ts';

/** 东亚行的单倍行高：(0.8 + 0.2) em × 1.3 */
const EA_LINE = SIZE_5 * 1.3;
/** 一行 10 个字 */
const TEN = '一二三四五六七八九十';

/** 版心 10 字宽（2100）× 3 行高（819） */
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

function opts(over: Partial<LayoutDocumentOptions> = {}): LayoutDocumentOptions {
  return { measurer: fakeMeasurer(), settings: DEFAULT_SETTINGS, ...over };
}

const paras = (doc: DocumentLayout, page: number): PlacedParagraph[] =>
  (doc.pages[page]?.blocks ?? []).filter((b): b is PlacedParagraph => b.kind === 'paragraph');

/** 每页各有几行（用来一眼看出分页点） */
const shape = (doc: DocumentLayout): number[] =>
  doc.pages.map((p) =>
    p.blocks.reduce((n, b) => n + (b.kind === 'paragraph' ? b.lines.length : b.rows.length), 0),
  );

describe('页面几何', () => {
  it('版心 = 纸张减页边距', () => {
    const g = pageGeometry(sect());
    expect(g.content).toEqual({ x: 600, y: 600, width: 2100, height: 819 });
  });

  it('装订线默认从左边扣，不是往纸外扩', () => {
    const g = pageGeometry(sect({ margin: { ...sect().margin, gutter: 300 } }));
    expect(g.content.x).toBe(900);
    expect(g.content.width).toBe(1800);
  });

  it('gutterAtTop 时装订线改从上边扣', () => {
    const props = sect({ margin: { ...sect().margin, gutter: 300 } });
    const g = pageGeometry(props, { settings: { ...DEFAULT_SETTINGS, gutterAtTop: true } });
    expect(g.content.x).toBe(600);
    expect(g.content.y).toBe(900);
    expect(g.content.height).toBe(519);
  });
});

describe('行摞起来', () => {
  it('行的 y 是前面各行行高的累加，段前间距把整段往下推', () => {
    const doc = layoutDocument(
      body([para([run(TEN + TEN)], { spacing: { ...paraProps().spacing, before: 100 } })]),
      opts(),
    );
    const p = paras(doc, 0)[0];
    expect(p?.y).toBe(100);
    expect(p?.lines.map((l) => l.y)).toEqual([100, 100 + EA_LINE]);
  });

  it('段后间距只影响下一段的起点，不算进本段的高度', () => {
    const a = para([run(TEN)], { spacing: { ...paraProps().spacing, after: 60 } });
    const b = para([run(TEN)]);
    const doc = layoutDocument(body([a, b]), opts());
    expect(paras(doc, 0)[1]?.y).toBe(EA_LINE + 60);
  });

  it('第一页的第一行贴着版心顶 —— 基线的绝对 y 靠 content.y + y + baseline 拼出来', () => {
    const doc = layoutDocument(body([para([run(TEN)])]), opts());
    const line = paras(doc, 0)[0]?.lines[0];
    expect(line?.y).toBe(0);
    expect(doc.pages[0]?.geometry.content.y).toBe(600);
    expect(line?.line.baseline).toBeGreaterThan(0);
  });
});

describe('自动分页', () => {
  it('放不下的行流到下一页，行下标接得上，first / last 标出段落被拆开', () => {
    const doc = layoutDocument(body([para([run(TEN.repeat(5))], { widowControl: false })]), opts());
    expect(shape(doc)).toEqual([3, 2]);

    const first = paras(doc, 0)[0];
    const second = paras(doc, 1)[0];
    expect(first?.id).toBe(second?.id); // 同一段，两片
    expect(first).toMatchObject({ first: true, last: false });
    expect(second).toMatchObject({ first: false, last: true });
    expect(first?.lines.map((l) => l.index)).toEqual([0, 1, 2]);
    expect(second?.lines.map((l) => l.index)).toEqual([3, 4]);
    // 续页从版心顶重新起算
    expect(second?.lines[0]?.y).toBe(0);
  });

  it('页码从 1 起、逐页加一，页序与页码是两个字段', () => {
    const doc = layoutDocument(body([para([run(TEN.repeat(5))], { widowControl: false })]), opts());
    expect(doc.pages.map((p) => [p.index, p.number])).toEqual([
      [0, 1],
      [1, 2],
    ]);
  });

  it('一行都放不下的版心也不会死循环 —— 硬塞一行并溢出', () => {
    const tiny = sect({ page: { width: 3300, height: 1300, orientation: 'portrait' } }); // 版心高 100
    const doc = layoutDocument(body([para([run(TEN.repeat(2))], { widowControl: false })], tiny), opts());
    expect(shape(doc)).toEqual([1, 1]);
  });
});

describe('孤行寡行（widowControl）', () => {
  it('下一页只接得走 1 行时，改成推 2 行过去', () => {
    // 4 行的段落：老实排是 3 + 1，寡行控制把它改成 2 + 2
    const doc = layoutDocument(body([para([run(TEN.repeat(4))])]), opts());
    expect(shape(doc)).toEqual([2, 2]);
  });

  it('本页只留得下 1 行时，整段挪到下一页', () => {
    // 前面一段占掉 2 行，后面 3 行的段落在本页只放得下 1 行 —— 孤行，整段走
    const doc = layoutDocument(body([para([run(TEN.repeat(2))]), para([run(TEN.repeat(3))])]), opts());
    expect(shape(doc)).toEqual([2, 3]);
  });

  it('关掉 widowControl 就老实排，页底能塞几行塞几行', () => {
    const doc = layoutDocument(body([para([run(TEN.repeat(4))], { widowControl: false })]), opts());
    expect(shape(doc)).toEqual([3, 1]);
  });

  it('两行的段落不受影响 —— 它本来就不该被拆', () => {
    const doc = layoutDocument(body([para([run(TEN)]), para([run(TEN.repeat(2))])]), opts());
    expect(shape(doc)).toEqual([3]);
  });
});

describe('keepLines / keepNext / pageBreakBefore', () => {
  it('keepLines：放不下就整段挪到下一页，不拆', () => {
    const doc = layoutDocument(
      body([para([run(TEN)]), para([run(TEN.repeat(3))], { keepLines: true })]),
      opts(),
    );
    expect(shape(doc)).toEqual([1, 3]);
  });

  it('keepNext：本段末行与下一段首行必须同页 —— 拆的是本段，不是整段推走', () => {
    // 4 行的 keepNext 段落 + 1 行的下一段。末行要留出接缝，于是第 1 页只放 2 行
    const doc = layoutDocument(
      body([para([run(TEN.repeat(4))], { keepNext: true }), para([run(TEN)])]),
      opts(),
    );
    expect(shape(doc)).toEqual([2, 3]);
    expect(paras(doc, 1).map((p) => p.lines.length)).toEqual([2, 1]);
  });

  it('keepNext 串成链：中间那段只有一行时，三段的接缝一起走', () => {
    const doc = layoutDocument(
      body([
        para([run(TEN.repeat(3))], { keepNext: true }),
        para([run(TEN)], { keepNext: true }),
        para([run(TEN)]),
      ]),
      opts(),
    );
    // 第一段的末行 + 第二段 + 第三段首行 = 3 行，正好一页；第一段前 2 行留在第 1 页
    expect(shape(doc)).toEqual([2, 3]);
  });

  it('pageBreakBefore：本段从新页开始；已经在页首就不再空跑一页', () => {
    const doc = layoutDocument(body([para([run(TEN)]), para([run(TEN)], { pageBreakBefore: true })]), opts());
    expect(shape(doc)).toEqual([1, 1]);

    const alone = layoutDocument(body([para([run(TEN)], { pageBreakBefore: true })]), opts());
    expect(alone.pages).toHaveLength(1);
  });
});

describe('硬分页符', () => {
  it('w:br type=page 在哪一行就断在哪一行之后', () => {
    const p = para([run(TEN), runOf([{ kind: 'break', breakType: 'page' }]), run(TEN)]);
    const doc = layoutDocument(body([p]), opts());
    expect(shape(doc)).toEqual([1, 1]);
  });

  it('段末的硬分页符不会凭空多出一张空页 —— 页是惰性开的', () => {
    const p = para([run(TEN), runOf([{ kind: 'break', breakType: 'page' }])]);
    const doc = layoutDocument(body([p]), opts());
    expect(doc.pages).toHaveLength(1);
  });
});

describe('分节', () => {
  const two = (a: SectionProps, b: SectionProps): ResolvedBody => ({
    sections: [
      { id: 's0', props: a, blocks: [para([run(TEN)])] },
      { id: 's1', props: b, blocks: [para([run(TEN)])] },
    ],
  });

  it('nextPage：新节另起一页', () => {
    const doc = layoutDocument(two(sect(), sect({ type: 'nextPage' })), opts());
    expect(shape(doc)).toEqual([1, 1]);
    expect(doc.pages.map((p) => p.sectionIndex)).toEqual([0, 1]);
  });

  it('continuous 且页面设置没变：接着上一节往下排，不换页', () => {
    const doc = layoutDocument(two(sect(), sect({ type: 'continuous' })), opts());
    expect(shape(doc)).toEqual([2]);
  });

  it('continuous 但页面设置变了：按换页处理并发诊断 —— 一页只有一个版心框', () => {
    const diagnostics = createDiagnosticSink();
    const wider = sect({ type: 'continuous', page: { width: 4000, height: 2019, orientation: 'portrait' } });
    const doc = layoutDocument(two(sect(), wider), opts({ diagnostics }));
    expect(shape(doc)).toEqual([1, 1]);
    expect(diagnostics.list().map((d) => d.code)).toEqual(['continuous-section-geometry-changed']);
  });

  it('oddPage：页码是偶数时补一张空页凑奇数', () => {
    const doc = layoutDocument(two(sect(), sect({ type: 'oddPage' })), opts());
    expect(doc.pages.map((p) => p.number)).toEqual([1, 2, 3]);
    expect(doc.pages[1]?.filler).toBe(true);
    expect(doc.pages[1]?.blocks).toEqual([]);
    expect(doc.pages[2]?.sectionIndex).toBe(1);
  });

  it('evenPage：页码已经是偶数就不补', () => {
    const first = sect();
    const doc = layoutDocument(
      {
        sections: [
          { id: 's0', props: first, blocks: [para([run(TEN.repeat(4))], { widowControl: false })] },
          { id: 's1', props: sect({ type: 'evenPage' }), blocks: [para([run(TEN)])] },
        ],
      },
      opts(),
    );
    // 第一节占 2 页（3 + 1 行），下一页是第 3 页 —— 奇数，要补一张才轮到偶数
    expect(doc.pages.map((p) => [p.number, p.filler === true])).toEqual([
      [1, false],
      [2, false],
      [3, true],
      [4, false],
    ]);
  });

  it('pgNumType 让页码重新起算，物理页序照旧', () => {
    const doc = layoutDocument(two(sect(), sect({ pageNumStart: 1 })), opts());
    expect(doc.pages.map((p) => [p.index, p.number])).toEqual([
      [0, 1],
      [1, 1],
    ]);
  });

  it('空节也占一页，页码序列不断档', () => {
    const doc = layoutDocument(
      {
        sections: [
          { id: 's0', props: sect(), blocks: [para([run(TEN)])] },
          { id: 's1', props: sect(), blocks: [] },
          { id: 's2', props: sect(), blocks: [para([run(TEN)])] },
        ],
      },
      opts(),
    );
    expect(doc.pages.map((p) => p.sectionIndex)).toEqual([0, 1, 2]);
  });

  it('多栏排版发诊断 —— 按单栏排出来的行长本来就是错的', () => {
    const diagnostics = createDiagnosticSink();
    layoutDocument(body([para([run(TEN)])], sect({ columns: 2 })), opts({ diagnostics }));
    expect(diagnostics.list()[0]?.code).toBe('multi-column-unsupported');
  });

  it('一份空文档也有一页', () => {
    expect(layoutDocument({ sections: [] }, opts()).pages).toHaveLength(1);
  });
});

describe('表格跨页', () => {
  const cellOf = (text: string) => cell([para([run(text)])]);
  const rows = (n: number, over: Parameters<typeof row>[1] = {}) =>
    Array.from({ length: n }, () => row([cellOf(TEN.slice(0, 4))], over));

  it('按行拆页：一行放不下就整行挪到下一页', () => {
    const t = table([2100], rows(5));
    const doc = layoutDocument(body([t]), opts());
    expect(shape(doc)).toEqual([3, 2]);

    const first = doc.pages[0]?.blocks[0] as PlacedTable;
    const second = doc.pages[1]?.blocks[0] as PlacedTable;
    expect(first.id).toBe(second.id);
    expect(first).toMatchObject({ first: true, last: false });
    expect(second).toMatchObject({ first: false, last: true });
    expect(second.rows.map((r) => r.index)).toEqual([3, 4]);
    expect(second.rows[0]?.y).toBe(0);
  });

  it('tblHeader 的表头在续页顶部重复，且**占掉**续页的一行高度', () => {
    const header = row([cellOf('表头')], { props: rowProps({ header: true }) });
    const t = table([2100], [header, ...rows(4)]);
    const doc = layoutDocument(body([t]), opts());

    const second = doc.pages[1]?.blocks[0] as PlacedTable;
    expect(second.rows.map((r) => [r.index, r.repeated === true])).toEqual([
      [0, true], // 重复出来的表头
      [3, false],
      [4, false],
    ]);
    // 第 1 页：表头 + 2 行正文；第 2 页：表头（重复）+ 2 行正文
    expect(shape(doc)).toEqual([3, 3]);
  });

  it('整张表都是表头行时不当表头处理 —— 否则续页会永远在重复表头', () => {
    const t = table([2100], rows(5, { props: rowProps({ header: true }) }));
    const doc = layoutDocument(body([t]), opts());
    expect(shape(doc)).toEqual([3, 2]);
  });
});

describe('结构化克隆', () => {
  it('整份分页结果可以过 Worker 边界', () => {
    const doc = layoutDocument(
      body([para([run(TEN.repeat(4))]), table([2100], [row([cell([para([run('甲')])])])])]),
      opts(),
    );
    expect(structuredClone(doc)).toEqual(doc);
  });
});
