/**
 * 布局索引：点 ↔ 位置的两个方向。
 *
 * 版心仍旧造成「一行 10 个字、一页 3 行」（与 page.test.ts 同一套），于是每个字的
 * 左边缘都是 600 + 210k 这种能心算的数 —— 断言里写的就是那个数，错了一眼看得出
 * 是差了一个字还是差了一行。
 *
 * 这里测的是**映射**，不是排版：字摆在哪由已经标定完的那几层负责（真值在
 * fixture.test.ts / page-fixture.test.ts …），索引只保证「画在那儿的字，点它能点中」。
 */
import type { NodeId, ResolvedBlock, ResolvedBody, SectionProps } from '@uw/model';
import { DEFAULT_SECTION_PROPS, DEFAULT_SETTINGS } from '@uw/model';
import { describe, expect, it } from 'vitest';
import { buildLayoutIndex } from './layout-index.ts';
import type { LayoutDocumentOptions } from './page.ts';
import { layoutDocument } from './page.ts';
import {
  cell,
  fakeMeasurer,
  NO_GRID,
  numberLabel,
  para,
  row,
  rowProps,
  run,
  runOf,
  SIZE_5,
  table,
} from './test-fixtures.ts';

/** 一行 10 个字 */
const TEN = '一二三四五六七八九十';
/** 东亚行的单倍行高：(0.8 + 0.2) em × 1.3 */
const EA_LINE = SIZE_5 * 1.3;
/** 版心左上角 */
const CX = 600;
const CY = 600;

function sect(over: Partial<SectionProps> = {}): SectionProps {
  return {
    ...structuredClone(DEFAULT_SECTION_PROPS),
    page: { width: 3300, height: 2019, orientation: 'portrait' },
    margin: { top: 600, right: 600, bottom: 600, left: 600, header: 300, footer: 300, gutter: 0 },
    docGrid: NO_GRID,
    ...over,
  };
}

function body(blocks: ResolvedBlock[], props: SectionProps = sect()): ResolvedBody {
  return { sections: [{ id: 's0', props, blocks }] };
}

function index(blocks: ResolvedBlock[], over: Partial<LayoutDocumentOptions> = {}, props = sect()) {
  return buildLayoutIndex(
    layoutDocument(body(blocks, props), { measurer: fakeMeasurer(), settings: DEFAULT_SETTINGS, ...over }),
  );
}

describe('positionAt · 点 → 位置', () => {
  it('点在第几个字上就给第几个偏移，跨行也对', () => {
    const r = run(TEN + TEN);
    const idx = index([para([r])]);

    // 第 1 行第 4 个字（下标 3）的字面上：左边缘 600 + 3 × 210 = 1230
    expect(idx.positionAt({ page: 0, x: 1230 + 50, y: CY + 10 })).toEqual({
      nodeId: r.id,
      contentIndex: 0,
      offset: 3,
    });
    // 第 2 行同一个位置，偏移多一整行
    expect(idx.positionAt({ page: 0, x: 1230 + 50, y: CY + EA_LINE + 10 })).toEqual({
      nodeId: r.id,
      contentIndex: 0,
      offset: 13,
    });
  });

  it('点落在字的右半边归下一个字缝 —— 光标该在字的后面', () => {
    const r = run(TEN);
    const idx = index([para([r])]);
    expect(idx.positionAt({ page: 0, x: CX + 200, y: CY + 10 })?.offset).toBe(1);
  });

  it('点在行外（页边距上）也答得出来 —— 取最近的那一行', () => {
    const r = run(TEN + TEN);
    const idx = index([para([r])]);
    // 纸的左上角：第 1 行的行首
    expect(idx.positionAt({ page: 0, x: 0, y: 0 })).toEqual({ nodeId: r.id, contentIndex: 0, offset: 0 });
    // 纸的右下角：最后一行的行尾（第 20 个字之后）
    expect(idx.positionAt({ page: 0, x: 3300, y: 2019 })).toEqual({
      nodeId: r.id,
      contentIndex: 0,
      offset: 20,
    });
  });

  it('同一个 run 里的两个 `w:t` 是两个位置 —— 片段在内容片的边界上切开', () => {
    const r = runOf([
      { kind: 'text', text: '甲' },
      { kind: 'text', text: '乙' },
    ]);
    const idx = index([para([r])]);
    const line = idx.lines[0]?.line;
    expect(line?.fragments.map((f) => [f.contentIndex, f.offset, f.text])).toEqual([
      [0, 0, '甲'],
      [1, 0, '乙'],
    ]);
    expect(idx.positionAt({ page: 0, x: CX + SIZE_5 + 50, y: CY + 10 })).toEqual({
      nodeId: r.id,
      contentIndex: 1,
      offset: 0,
    });
  });

  it('编号不可定位：点在编号上给的是正文的第一个字', () => {
    const r = run('正文');
    const idx = index([para([r], { numbering: { numId: 1, level: 0, label: numberLabel('一、') } })]);
    // 编号「一、」占前两个字宽，点在它正中间
    expect(idx.positionAt({ page: 0, x: CX + SIZE_5, y: CY + 10 })).toEqual({
      nodeId: r.id,
      contentIndex: 0,
      offset: 0,
    });
  });

  it('域结果不可定位：它不是文件里那串字符', () => {
    const f = run('9');
    const r = run('页');
    const idx = index([para([f, r])], { fieldValues: new Map<NodeId, string>([[f.id, '123']]) });
    // 域结果占了前 1.5 个字宽（3 个 ASCII 数字），点在它身上落到后面那个真实的字上
    expect(idx.positionAt({ page: 0, x: CX + 100, y: CY + 10 })?.nodeId).toBe(r.id);
  });
});

describe('rectsOf · 位置 → 矩形', () => {
  it('跨行的 range 一行一个矩形', () => {
    const r = run(TEN + TEN);
    const idx = index([para([r])]);
    const rects = idx.rectsOf({
      start: { nodeId: r.id, contentIndex: 0, offset: 5 },
      end: { nodeId: r.id, contentIndex: 0, offset: 15 },
    });
    expect(rects).toEqual([
      { page: 0, x: CX + 5 * SIZE_5, y: CY, width: 5 * SIZE_5, height: EA_LINE },
      { page: 0, x: CX, y: CY + EA_LINE, width: 5 * SIZE_5, height: EA_LINE },
    ]);
  });

  it('两端给反了也算得出来 —— 先规范化', () => {
    const r = run(TEN);
    const idx = index([para([r])]);
    const forward = idx.rectsOf({
      start: { nodeId: r.id, contentIndex: 0, offset: 2 },
      end: { nodeId: r.id, contentIndex: 0, offset: 4 },
    });
    const backward = idx.rectsOf({
      start: { nodeId: r.id, contentIndex: 0, offset: 4 },
      end: { nodeId: r.id, contentIndex: 0, offset: 2 },
    });
    expect(backward).toEqual(forward);
  });

  it('一行里相邻的片段并成一个矩形（同一行两个 run 不该给两块）', () => {
    const a = run('甲乙');
    const b = run('丙丁');
    const idx = index([para([a, b])]);
    const rects = idx.rectsOf({
      start: { nodeId: a.id, contentIndex: 0, offset: 0 },
      end: { nodeId: b.id, contentIndex: 0, offset: 2 },
    });
    expect(rects).toEqual([{ page: 0, x: CX, y: CY, width: 4 * SIZE_5, height: EA_LINE }]);
  });

  it('整段选中时编号不在里面 —— 它不在 document.xml 里，复制不出来', () => {
    const r = run('正文');
    const idx = index([para([r], { numbering: { numId: 1, level: 0, label: numberLabel('一、') } })]);
    const rects = idx.rectsOf({
      start: { nodeId: r.id, contentIndex: 0, offset: 0 },
      end: { nodeId: r.id, contentIndex: 0, offset: 2 },
    });
    // 正文被制表位顶到默认制表位（720）上，矩形从那儿起 —— 没有把编号那两个字圈进去
    expect(rects).toHaveLength(1);
    expect(rects[0]?.x).toBe(CX + DEFAULT_SETTINGS.defaultTabStop);
  });

  it('run 没排出来时给空 —— 空 run 上没有任何一个字形', () => {
    const r = run(TEN);
    const idx = index([para([r])]);
    expect(
      idx.rectsOf({
        start: { nodeId: 'not-laid-out', contentIndex: 0, offset: 0 },
        end: { nodeId: r.id, contentIndex: 0, offset: 1 },
      }),
    ).toEqual([]);
  });
});

describe('caretRect · 位置 → 光标', () => {
  it('与 positionAt 互逆', () => {
    const r = run(TEN + TEN);
    const idx = index([para([r])]);
    for (const offset of [0, 7, 10, 19]) {
      const rect = idx.caretRect({ nodeId: r.id, contentIndex: 0, offset });
      expect(rect).toBeDefined();
      const back = idx.positionAt({
        page: rect?.page as number,
        x: (rect?.x as number) + 1,
        y: (rect?.y as number) + 10,
      });
      expect(back).toEqual({ nodeId: r.id, contentIndex: 0, offset });
    }
  });

  it('断行处的光标落在下一行行首，不是上一行行尾', () => {
    const r = run(TEN + TEN);
    const idx = index([para([r])]);
    expect(idx.caretRect({ nodeId: r.id, contentIndex: 0, offset: 10 })).toEqual({
      page: 0,
      x: CX,
      y: CY + EA_LINE,
      width: 0,
      height: EA_LINE,
    });
  });

  it('段末之后的那一缝在最后一个字的右边', () => {
    const r = run('甲乙');
    const idx = index([para([r])]);
    expect(idx.caretRect({ nodeId: r.id, contentIndex: 0, offset: 2 })?.x).toBe(CX + 2 * SIZE_5);
  });
});

describe('compare · 文档序', () => {
  it('按 run 的出现顺序，再按内容片段与偏移', () => {
    const a = run('甲');
    const b = run('乙');
    const idx = index([para([a]), para([b])]);
    const pa = { nodeId: a.id, contentIndex: 0, offset: 0 };
    expect(idx.compare(pa, { nodeId: b.id, contentIndex: 0, offset: 0 })).toBe(-1);
    expect(idx.compare(pa, pa)).toBe(0);
    expect(idx.compare({ nodeId: a.id, contentIndex: 0, offset: 1 }, pa)).toBe(1);
  });

  it('排不出来的 run 给 undefined，不给 0', () => {
    const a = run('甲');
    const idx = index([para([a])]);
    const pa = { nodeId: a.id, contentIndex: 0, offset: 0 };
    expect(idx.compare({ nodeId: 'ghost', contentIndex: 0, offset: 0 }, pa)).toBeUndefined();
    expect(idx.compare(pa, { nodeId: 'ghost', contentIndex: 0, offset: 0 })).toBeUndefined();
  });
});

describe('容器 · 表格与页眉页脚', () => {
  it('格内的行带着格子的原点（表格缩进 + 格子 x + 左边距）', () => {
    const left = run('甲');
    const right = run('乙');
    const idx = index([table([1050, 1050], [row([cell([para([left])]), cell([para([right])])])])]);

    // 右格：格子左边 1050 + 默认左边距 108
    expect(idx.positionAt({ page: 0, x: CX + 1050 + 108 + 50, y: CY + 10 })?.nodeId).toBe(right.id);
    expect(idx.positionAt({ page: 0, x: CX + 108 + 50, y: CY + 10 })?.nodeId).toBe(left.id);
    // 两格的行在同一个 y 上 —— 命中先比纵向再比横向，靠横向分的胜负
    expect(idx.lines.map((l) => l.top)).toEqual([CY, CY]);
  });

  it('跨页重复的表头行标出来，且指回文档里那一份', () => {
    const head = run('头');
    const rows = [row([cell([para([head])])], { props: rowProps({ header: true }) })];
    for (let i = 0; i < 5; i++) rows.push(row([cell([para([run(`第${i}行`)])])]));
    const idx = index([table([2100], rows)]);

    // 表头 + 5 行数据、一页 3 行 → 第 2 / 3 页顶上各重复一遍表头
    const repeated = idx.lines.filter((l) => l.repeated);
    expect(repeated.map((l) => l.page)).toEqual([1, 2]);
    // 点它给的是文档里那一份表头（Word 也是把光标放进去），不是第二页上凭空多出来的东西
    const p = idx.positionAt({ page: 1, x: CX + 150, y: (repeated[0]?.top as number) + 10 });
    expect(p?.nodeId).toBe(head.id);
  });

  it('页眉里的行相对纸左上角，且标着 frame', () => {
    const h = run('眉');
    const idx = index(
      [para([run('正文')])],
      {
        headerFooters: { h: { resolved: [para([h])] } },
        // 页眉框顶 = w:header = 300（实测，见 HEADER_RULES）
      },
      sect({ headers: [{ type: 'default', relId: 'h' }] }),
    );

    const line = idx.lines[0];
    expect(line?.frame).toBe('header');
    expect(line?.top).toBe(300);
    expect(idx.positionAt({ page: 0, x: CX + 50, y: 300 + 10 })?.nodeId).toBe(h.id);
  });
});
