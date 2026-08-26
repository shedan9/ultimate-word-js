/**
 * 画法的单测。全部用**手搭的** `PageLayout`，不跑真实文档：这里要验的是
 * 「布局给的 twips 有没有被原封不动地翻译成 pt」，掺进真实文档的排版结果反而
 * 让失败信息变成「差了 0.3pt，不知道是排版的锅还是画法的锅」。
 * 端到端那一半在 fixture.test.ts，那里比的是真值。
 */
import type {
  CellLayout,
  FragmentStyle,
  LineFragment,
  LineLayout,
  PageLayout,
  PlacedRow,
  PlacedTable,
  RowLayout,
} from '@uw/layout';
import { describe, expect, it } from 'vitest';
import { buildPage } from './paint.ts';
import type { RElement } from './tree.ts';
import { serialize } from './tree.ts';

const STYLE: FragmentStyle = {
  bold: false,
  italic: false,
  color: 'auto',
  underline: 'none',
  strike: false,
  doubleStrike: false,
  vertAlign: 'baseline',
  position: 0,
  scale: 100,
};

/** A4 纵向 + 2.54cm 页边距，与 fixture 的量级一致，算出来的 pt 一眼能认 */
const GEOMETRY = {
  width: 11906,
  height: 16838,
  content: { x: 1440, y: 1440, width: 9026, height: 13958 },
};

function frag(over: Partial<LineFragment> = {}): LineFragment {
  return {
    runId: 'r1',
    font: '仿宋',
    fontSize: 320,
    script: 'eastAsia',
    style: STYLE,
    text: '甲乙',
    x: 0,
    width: 640,
    glyphX: [0, 320],
    ...over,
  };
}

function line(over: Partial<LineLayout> = {}): LineLayout {
  return {
    start: 0,
    end: 2,
    x: 0,
    width: 640,
    height: 480,
    baseline: 380,
    natural: 480,
    fragments: [frag()],
    leaders: [],
    isLast: true,
    ...over,
  };
}

function page(blocks: PageLayout['blocks']): PageLayout {
  return { index: 0, number: 1, sectionIndex: 0, geometry: GEOMETRY, blocks };
}

function paragraphPage(l: LineLayout = line()): PageLayout {
  return page([
    { kind: 'paragraph', id: 'p1', y: 0, lines: [{ index: 0, y: 0, line: l }], first: true, last: true },
  ]);
}

/** 深度优先收集某个标签的所有元素 —— 断言时只关心「有几个、什么属性」 */
function collect(node: RElement, tag: string, out: RElement[] = []): RElement[] {
  if (node.tag === tag) out.push(node);
  for (const c of node.children) collect(c, tag, out);
  return out;
}

describe('页', () => {
  it('viewBox 的单位是 pt，width / height 才是 px', () => {
    const svg = buildPage(paragraphPage());
    // 11906 twips = 595.3pt = 793.733px（96dpi）
    expect(svg.attrs.viewBox).toBe('0 0 595.3 841.9');
    expect(svg.attrs.width).toBe('793.733px');
    expect(svg.attrs.height).toBe('1122.533px');
  });

  it('缩放只改 width / height，viewBox 一个字不动', () => {
    const a = buildPage(paragraphPage());
    const b = buildPage(paragraphPage(), { zoom: 2 });
    expect(b.attrs.viewBox).toBe(a.attrs.viewBox);
    expect(b.attrs.width).toBe('1587.467px');
  });

  it('版心原点落在 content 的左上角', () => {
    const svg = buildPage(paragraphPage());
    const g = collect(svg, 'g').find((n) => n.attrs.class === 'uw-content');
    expect(g?.attrs.transform).toBe('translate(72 72)');
  });
});

describe('页眉页脚', () => {
  const frame = (kind: 'header' | 'footer', y: number) => ({
    kind,
    relId: `r-${kind}`,
    x: 1440,
    y,
    width: 9026,
    height: 480,
    blocks: [
      {
        kind: 'paragraph' as const,
        id: `p-${kind}`,
        y: 0,
        lines: [{ index: 0, y: 0, line: line({ fragments: [frag({ text: kind })] }) }],
        first: true,
        last: true,
      },
    ],
  });

  const withFrames = (): PageLayout => ({
    ...paragraphPage(),
    header: frame('header', 851),
    footer: frame('footer', 15507),
  });

  it('框的坐标相对**纸**左上角，与版心那个 g 平级 —— 套进版心会平白多偏一个上边距', () => {
    const svg = buildPage(withFrames());
    const g = collect(svg, 'g');
    // 851 twips = 42.55pt、15507 twips = 775.35pt
    expect(g.find((n) => n.attrs.class === 'uw-header')?.attrs.transform).toBe('translate(72 42.55)');
    expect(g.find((n) => n.attrs.class === 'uw-footer')?.attrs.transform).toBe('translate(72 775.35)');
  });

  it('页眉在正文之前画、页脚在之后 —— 重叠时正文压在页眉上更容易看出是哪儿排错了', () => {
    const svg = buildPage(withFrames());
    const order = svg.children.map((c) => c.attrs.class).filter((c) => c !== undefined);
    expect(order).toEqual(['uw-page-bg', 'uw-header', 'uw-content', 'uw-footer']);
  });

  it('没有页眉页脚的页与从前一模一样', () => {
    const svg = buildPage(paragraphPage());
    expect(collect(svg, 'g').some((n) => n.attrs.class === 'uw-header')).toBe(false);
  });
});

describe('文字片段', () => {
  it('基线 = 版心顶 + 行顶 + 行内基线，逐字 x 直接进 x 列表', () => {
    const svg = buildPage(paragraphPage());
    const t = collect(svg, 'text')[0] as RElement;
    // 行顶 0 + 行内基线 380 = 380 twips = 19pt。版心顶那 72pt 由外层 <g> 的 translate 带着 ——
    // 页面绝对基线 = 72 + 19 = 91pt，与 truth.json 里那一列是同一个数
    expect(t.attrs.y).toBe('19');
    expect(t.attrs.x).toBe('0 16');
    expect(t.attrs['font-size']).toBe('16');
    expect(t.text).toBe('甲乙');
    expect(t.attrs['xml:space']).toBe('preserve');
  });

  it('w:position 抬基线，上标再抬一次 —— 两者可叠加', () => {
    const raised = { ...STYLE, position: 20, vertAlign: 'superscript' as const };
    const svg = buildPage(paragraphPage(line({ fragments: [frag({ style: raised })] })));
    const t = collect(svg, 'text')[0] as RElement;
    // 380 - 20 = 360 twips = 18pt，再减 16pt × 0.45（SUPERSCRIPT_RAISE_EM）
    expect(Number(t.attrs.y)).toBeCloseTo(18 - 16 * 0.45, 3);
  });

  it('w:w 横向缩放用 transform，x 同时除回去', () => {
    const squeezed = { ...STYLE, scale: 50 };
    const svg = buildPage(paragraphPage(line({ fragments: [frag({ style: squeezed, glyphX: [0, 160] })] })));
    const t = collect(svg, 'text')[0] as RElement;
    expect(t.attrs.transform).toBe('scale(0.5 1)');
    // 第二个字实际落在 160 twips = 8pt，坐标系压扁一半后要写 16
    expect(t.attrs.x).toBe('0 16');
  });

  it('编号片段带标记 —— 可选文本层与复制要靠它跳过', () => {
    const svg = buildPage(paragraphPage(line({ fragments: [frag({ numbering: true, text: '一、' })] })));
    const t = collect(svg, 'text')[0] as RElement;
    expect(t.attrs['data-numbering']).toBe('1');
  });
});

describe('装饰', () => {
  it('下划线画在文字之前 —— 画在后面会盖住字的下半截', () => {
    const style = { ...STYLE, underline: 'single' };
    const svg = buildPage(paragraphPage(line({ fragments: [frag({ style })] })));
    const g = collect(svg, 'g').find((n) => n.attrs.class === 'uw-para') as RElement;
    expect(g.children.map((c) => c.tag)).toEqual(['rect', 'text']);
    expect(Number((g.children[0] as RElement).attrs.y)).toBeGreaterThan(19);
  });

  it('双删除线画两条，单删除线画一条', () => {
    const one = buildPage(paragraphPage(line({ fragments: [frag({ style: { ...STYLE, strike: true } })] })));
    const two = buildPage(
      paragraphPage(line({ fragments: [frag({ style: { ...STYLE, doubleStrike: true } })] })),
    );
    expect(collect(one, 'rect').filter((r) => r.attrs.class === 'uw-strike')).toHaveLength(1);
    expect(collect(two, 'rect').filter((r) => r.attrs.class === 'uw-strike')).toHaveLength(2);
  });

  it('没有装饰时一个多余元素都不出', () => {
    const svg = buildPage(paragraphPage());
    const g = collect(svg, 'g').find((n) => n.attrs.class === 'uw-para') as RElement;
    expect(g.children).toHaveLength(1);
  });
});

// ── 表格 ──────────────────────────────────────────────────────────────────────

const SINGLE = { style: 'single', size: 10, space: 0, color: '000000' };

function cell(over: Partial<CellLayout> = {}): CellLayout {
  return {
    cellId: 'c1',
    col: 0,
    span: 1,
    x: 0,
    width: 4000,
    contentWidth: 3784,
    paddingLeft: 108,
    paddingRight: 108,
    paddingTop: 0,
    paddingBottom: 0,
    verticalAlign: 'top',
    shading: undefined,
    vMerge: 'none',
    borders: { top: [], bottom: [], left: undefined, right: undefined, tl2br: undefined, tr2bl: undefined },
    blocks: [],
    contentHeight: 480,
    ...over,
  };
}

function tablePage(row: RowLayout, columns: number[]): PageLayout {
  const table: PlacedTable = {
    kind: 'table',
    id: 't1',
    x: 0,
    y: 0,
    width: columns.reduce((a, b) => a + b, 0),
    columns,
    rows: [{ index: 0, y: 0, height: row.height, row }],
    first: true,
    last: true,
  };
  return page([table]);
}

describe('表格', () => {
  it('共享的那条格线只画一次', () => {
    const left = cell({ cellId: 'a', borders: { ...cell().borders, right: SINGLE } });
    const right = cell({
      cellId: 'b',
      col: 1,
      x: 4000,
      width: 5000,
      borders: { ...cell().borders, left: SINGLE },
    });
    const svg = buildPage(
      tablePage({ rowId: 'r1', cells: [left, right], gridAbove: 0, height: 480 }, [4000, 5000]),
    );
    const lines = collect(svg, 'line');
    expect(lines).toHaveLength(1);
    expect(lines[0]?.attrs.x1).toBe('200');
    expect(lines[0]?.attrs.y2).toBe('24');
  });

  it('nil 与 none 不画，size 为 0 也不画', () => {
    const c = cell({
      borders: {
        ...cell().borders,
        left: { ...SINGLE, style: 'nil' },
        right: { ...SINGLE, size: 0 },
      },
    });
    const svg = buildPage(tablePage({ rowId: 'r1', cells: [c], gridAbove: 0, height: 480 }, [4000]));
    expect(collect(svg, 'line')).toHaveLength(0);
  });

  it('水平边按网格列分段 —— 上面一格跨两列、下面两格时那条线分两段', () => {
    const c = cell({
      width: 9000,
      span: 2,
      borders: {
        ...cell().borders,
        bottom: [
          { col: 0, span: 1, border: SINGLE },
          { col: 1, span: 1, border: undefined },
        ],
      },
    });
    const svg = buildPage(tablePage({ rowId: 'r1', cells: [c], gridAbove: 0, height: 480 }, [4000, 5000]));
    const lines = collect(svg, 'line');
    expect(lines).toHaveLength(1);
    expect([lines[0]?.attrs.x1, lines[0]?.attrs.x2]).toEqual(['0', '200']);
  });

  it('底纹取 fill 不取 color，clear 以外的图案也只铺纯色', () => {
    const c = cell({ shading: { pattern: 'clear', color: 'FF0000', fill: 'D9D9D9' } });
    const svg = buildPage(tablePage({ rowId: 'r1', cells: [c], gridAbove: 0, height: 480 }, [4000]));
    const bg = collect(svg, 'rect').find((r) => r.attrs.class === 'uw-cell-bg');
    expect(bg?.attrs.fill).toBe('#d9d9d9');
  });

  it('auto / nil 的底纹不铺 —— 铺了会把页面背景压成白块', () => {
    const c = cell({ shading: { pattern: 'clear', color: 'auto', fill: 'auto' } });
    const svg = buildPage(tablePage({ rowId: 'r1', cells: [c], gridAbove: 0, height: 480 }, [4000]));
    expect(collect(svg, 'rect').filter((r) => r.attrs.class === 'uw-cell-bg')).toHaveLength(0);
  });

  it('w:vAlign=center 把格内内容摞到竖直中间', () => {
    const para = {
      kind: 'paragraph' as const,
      layout: { paragraphId: 'p1', lines: [line()], spaceBefore: 0, spaceAfter: 0, contentWidth: 3784 },
    };
    const top = cell({ blocks: [para], paddingTop: 100, paddingBottom: 100 });
    const mid = cell({ ...top, verticalAlign: 'center' });
    const yOf = (c: CellLayout): number => {
      const svg = buildPage(tablePage({ rowId: 'r1', cells: [c], gridAbove: 0, height: 1480 }, [4000]));
      return Number((collect(svg, 'text')[0] as RElement).attrs.y);
    };
    // 行高 480、可用高 1480-200=1280，居中要往下挪 (1280-480)/2 = 400 twips = 20pt
    expect(yOf(mid) - yOf(top)).toBeCloseTo(20, 3);
  });

  it('拆行的接缝上不画横线，竖边照旧两片各画各的', () => {
    const c = cell({
      borders: {
        ...cell().borders,
        top: [{ col: 0, span: 1, border: SINGLE }],
        bottom: [{ col: 0, span: 1, border: SINGLE }],
        left: SINGLE,
      },
    });
    const row: RowLayout = { rowId: 'r1', cells: [c], gridAbove: 0, height: 480 };
    const slice = (over: Partial<PlacedRow>): RElement => {
      const t: PlacedTable = {
        kind: 'table',
        id: 't1',
        x: 0,
        y: 0,
        width: 4000,
        columns: [4000],
        rows: [{ index: 0, y: 0, height: 480, row, ...over }],
        first: true,
        last: true,
      };
      return buildPage(page([t]));
    };
    const horizontals = (svg: RElement): number =>
      collect(svg, 'line').filter((l) => l.attrs.y1 === l.attrs.y2).length;

    expect(horizontals(slice({}))).toBe(2); // 整行：上下两条都是真边界
    expect(horizontals(slice({ splitAfter: true }))).toBe(1); // 底是切口
    expect(horizontals(slice({ continued: true }))).toBe(1); // 顶是切口
    expect(horizontals(slice({ continued: true, splitAfter: true }))).toBe(0);
    // 竖边不受影响：哪一片都要画自己那一截
    expect(collect(slice({ continued: true, splitAfter: true }), 'line')).toHaveLength(1);
  });

  it('vMerge=continue 的格子不画内容，但格线照旧参与', () => {
    const para = {
      kind: 'paragraph' as const,
      layout: { paragraphId: 'p1', lines: [line()], spaceBefore: 0, spaceAfter: 0, contentWidth: 3784 },
    };
    const c = cell({ vMerge: 'continue', blocks: [para], borders: { ...cell().borders, left: SINGLE } });
    const svg = buildPage(tablePage({ rowId: 'r1', cells: [c], gridAbove: 0, height: 480 }, [4000]));
    expect(collect(svg, 'text')).toHaveLength(0);
    expect(collect(svg, 'line')).toHaveLength(1);
  });
});

describe('序列化', () => {
  it('整页能吐成一段合法的 SVG 标记', () => {
    const out = serialize(buildPage(paragraphPage()));
    expect(out.startsWith('<svg xmlns="http://www.w3.org/2000/svg"')).toBe(true);
    expect(out.endsWith('</svg>')).toBe(true);
    expect(out).toContain('>甲乙</text>');
  });
});
