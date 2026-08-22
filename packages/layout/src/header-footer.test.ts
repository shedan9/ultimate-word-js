/**
 * 页眉页脚：选哪一份、摆在哪儿、把版心挤成多高。
 *
 * 几何那三条的**真值**判据在 `header-fixture.test.ts`（三份 Word 样本 12 页逐行对）。
 * 这里测的是真值样本覆盖不到的那些格：合成的版心让「一页几行」可以手算，
 * 选择规则里 `w:titlePg` 没定义 first、跨节沿用上一节这类分支也只有合成样本造得出来。
 */
import type { NodeId, ResolvedBlock, ResolvedBody, SectionProps } from '@uw/model';
import { DEFAULT_SECTION_PROPS, DEFAULT_SETTINGS } from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { HeaderFooterSource } from './header-footer.ts';
import { contentWithHeaderFooter, frameOf, pickHeaderFooter } from './header-footer.ts';
import type { DocumentLayout, LayoutDocumentOptions, PageGeometry } from './page.ts';
import { layoutDocument, pageGeometry } from './page.ts';
import { fakeMeasurer, NO_GRID, para, run, SIZE_5 } from './test-fixtures.ts';

/** 一行 10 个字（2100 twips 宽），一页 3 行（行高 = 1.3 em = 273） */
const TEN = '一二三四五六七八九十';

function sect(over: Partial<SectionProps> = {}): SectionProps {
  return {
    ...structuredClone(DEFAULT_SECTION_PROPS),
    page: { width: 3300, height: 2019, orientation: 'portrait' },
    margin: { top: 600, right: 600, bottom: 600, left: 600, header: 300, footer: 300, gutter: 0 },
    docGrid: NO_GRID,
    ...over,
  };
}

describe('pickHeaderFooter · 这一页该用哪一份', () => {
  const ref = (type: 'default' | 'first' | 'even', relId: string) => ({ type, relId });
  const pick = (
    sections: SectionProps[],
    index: number,
    firstInSection: boolean,
    pageNumber: number,
    evenAndOdd = false,
  ) =>
    pickHeaderFooter(sections, index, 'header', firstInSection, pageNumber, {
      ...DEFAULT_SETTINGS,
      evenAndOddHeaders: evenAndOdd,
    })?.relId;

  it('平常就是 default', () => {
    const s = [sect({ headers: [ref('default', 'd')] })];
    expect(pick(s, 0, true, 1)).toBe('d');
    expect(pick(s, 0, false, 2)).toBe('d');
  });

  it('w:titlePg 只作用在本节的第一页上', () => {
    const s = [sect({ titlePage: true, headers: [ref('default', 'd'), ref('first', 'f')] })];
    expect(pick(s, 0, true, 1)).toBe('f');
    expect(pick(s, 0, false, 2)).toBe('d');
  });

  it('w:titlePg 开着却没定义 first 时，首页是**空的**而不是退回 default', () => {
    // 退回 default 会让每份带封面的公文首页平白多出一行页眉
    const s = [sect({ titlePage: true, headers: [ref('default', 'd')] })];
    expect(pick(s, 0, true, 1)).toBeUndefined();
    expect(pick(s, 0, false, 2)).toBe('d');
  });

  it('奇偶页看的是**显示页码**，不是物理页序', () => {
    const s = [sect({ headers: [ref('default', 'd'), ref('even', 'e')] })];
    expect(pick(s, 0, true, 1, true)).toBe('d');
    expect(pick(s, 0, false, 2, true)).toBe('e');
    // 设置没开时 even 那一份根本不参与
    expect(pick(s, 0, false, 2, false)).toBe('d');
  });

  it('首页与偶数页撞车时 titlePg 赢', () => {
    const s = [
      sect({ titlePage: true, headers: [ref('default', 'd'), ref('first', 'f'), ref('even', 'e')] }),
    ];
    // 页码从 2 起算：第一张纸既是「本节首页」又是「偶数页」
    expect(pick(s, 0, true, 2, true)).toBe('f');
  });

  it('本节没写这一类就沿用前一节的（「链接到上一节」在文件里就是不写）', () => {
    const s = [sect({ headers: [ref('default', 'd')] }), sect({ headers: [] })];
    expect(pick(s, 1, false, 2)).toBe('d');
  });

  it('一节都没定义就是没有页眉', () => {
    expect(pick([sect()], 0, true, 1)).toBeUndefined();
  });
});

describe('几何 · 页边距是最小值不是固定值', () => {
  const base = (): PageGeometry => pageGeometry(sect());
  const margin = sect().margin;

  it('放得下时正文一动不动', () => {
    const g = contentWithHeaderFooter(base(), margin, 200, 200);
    expect(g.content.y).toBe(600);
    expect(g.content.height).toBe(2019 - 600 - 600);
  });

  it('页眉长过上边距时把版心顶下去', () => {
    const g = contentWithHeaderFooter(base(), margin, 500, 0);
    expect(g.content.y).toBe(300 + 500);
    // 底不动，所以版心整个变矮
    expect(g.content.y + g.content.height).toBe(2019 - 600);
  });

  it('页脚长过下边距时把版心顶上来', () => {
    const g = contentWithHeaderFooter(base(), margin, 0, 900);
    expect(g.content.y).toBe(600);
    expect(g.content.y + g.content.height).toBe(2019 - 300 - 900);
  });

  it('页眉页脚长到把版心吃光时钳到 0，而不是给出负高度', () => {
    // 负高度会让每一页都「一行都放不下」，于是硬塞一行、无限翻页
    const g = contentWithHeaderFooter(base(), margin, 1200, 1200);
    expect(g.content.height).toBe(0);
  });

  it('页眉从纸顶往下长，页脚从纸底往上长', () => {
    expect(frameOf('header', base(), margin, 500).y).toBe(300);
    expect(frameOf('footer', base(), margin, 500).y).toBe(2019 - 300 - 500);
    // 横向与版心同宽 —— 左右页边距是共用的
    expect(frameOf('header', base(), margin, 500).width).toBe(2100);
  });
});

describe('分页 · 页眉页脚参与进来之后', () => {
  const source = (blocks: Record<string, ResolvedBlock[]>): HeaderFooterSource =>
    Object.fromEntries(Object.entries(blocks).map(([k, v]) => [k, { resolved: v }]));

  function body(props: SectionProps, blocks: ResolvedBlock[]): ResolvedBody {
    return { sections: [{ id: 's0', props, blocks }] };
  }

  function place(props: SectionProps, over: Partial<LayoutDocumentOptions> = {}): DocumentLayout {
    // 12 行的一段，页高够 3 行时是 4 页
    const blocks = [para([run(TEN.repeat(12))])];
    return layoutDocument(body(props, blocks), {
      measurer: fakeMeasurer(),
      settings: DEFAULT_SETTINGS,
      ...over,
    });
  }

  it('没传 headerFooters 时版心就是「纸减页边距」，与从前一模一样', () => {
    const doc = place(sect({ headers: [{ type: 'default', relId: 'h' }] }));
    expect(doc.pages[0]?.header).toBeUndefined();
    expect(doc.pages[0]?.geometry.content.y).toBe(600);
  });

  it('矮页眉画出来但不动版心', () => {
    const doc = place(sect({ headers: [{ type: 'default', relId: 'h' }] }), {
      headerFooters: source({ h: [para([run('眉')])] }),
    });
    expect(doc.pages[0]?.header?.relId).toBe('h');
    expect(doc.pages[0]?.header?.height).toBe(SIZE_5 * 1.3);
    expect(doc.pages[0]?.geometry.content.y).toBe(600);
    expect(doc.pages).toHaveLength(4);
  });

  it('高页眉顶开版心，一页因此少排一行、总页数变多', () => {
    const tall = [para([run('眉')]), para([run('眉')])];
    const doc = place(sect({ headers: [{ type: 'default', relId: 'h' }] }), {
      headerFooters: source({ h: tall }),
    });
    const h = SIZE_5 * 1.3 * 2;
    expect(doc.pages[0]?.geometry.content.y).toBe(300 + h);
    // 版心从 3 行缩到 2 行，12 行于是排成 6 页
    expect(doc.pages[0]?.blocks[0]?.kind === 'paragraph' && doc.pages[0].blocks[0].lines).toHaveLength(2);
    expect(doc.pages).toHaveLength(6);
  });

  it('内容是空的引用当没有页眉处理 —— 别为它白留一截高度', () => {
    const doc = place(sect({ headers: [{ type: 'default', relId: 'h' }] }), {
      headerFooters: source({ h: [] }),
    });
    expect(doc.pages[0]?.header).toBeUndefined();
  });

  it('静态页眉全文档共用同一份数据对象（几百页不复制几百遍）', () => {
    const doc = place(sect({ headers: [{ type: 'default', relId: 'h' }] }), {
      headerFooters: source({ h: [para([run('眉')])] }),
    });
    expect(doc.pages[0]?.header?.blocks).toBe(doc.pages[3]?.header?.blocks);
  });

  it('页脚里的 PAGE 一趟就是准的 —— 页码在开页那一刻就定了，不必等下一趟迭代', () => {
    const r = run('9');
    const fields = new Map<NodeId, { type: 'PAGE' }>([[r.id, { type: 'PAGE' }]]);
    const doc = place(sect({ footers: [{ type: 'default', relId: 'f' }] }), {
      headerFooters: source({ f: [para([r])] }),
      headerFields: { fields },
    });
    const texts = doc.pages.map((p) => {
      const b = p.footer?.blocks[0];
      return b?.kind === 'paragraph' ? b.lines[0]?.line.fragments.map((f) => f.text).join('') : undefined;
    });
    expect(texts).toEqual(['1', '2', '3', '4']);
  });

  it('NUMPAGES 在还不知道总页数的那一趟保留文件里的旧值', () => {
    const r = run('99');
    const fields = new Map<NodeId, { type: 'NUMPAGES' }>([[r.id, { type: 'NUMPAGES' }]]);
    const doc = place(sect({ footers: [{ type: 'default', relId: 'f' }] }), {
      headerFooters: source({ f: [para([r])] }),
      headerFields: { fields },
    });
    const b = doc.pages[0]?.footer?.blocks[0];
    expect(b?.kind === 'paragraph' && b.lines[0]?.line.fragments[0]?.text).toBe('99');
  });
});
