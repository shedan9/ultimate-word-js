/**
 * 域求值。
 *
 * 版心与 page.test.ts 同一套：一行 10 个字、一页 3 行，于是「域落在第几页」可以数着行写出来。
 * 这里测的是**规则**（算成几、格式怎么选、迭代到哪一趟停），页码本身准不准由分页那边
 * 的真值断言兜着（`page-fixture.test.ts` 的 50 页）。
 */
import { createDiagnosticSink } from '@uw/core';
import type { FieldRegion, NodeId, ResolvedBlock, ResolvedBody, ResolvedRun, SectionProps } from '@uw/model';
import { DEFAULT_SECTION_PROPS, DEFAULT_SETTINGS, parseFieldInstruction } from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { LayoutDocumentWithFieldsOptions } from './fields.ts';
import { layoutDocumentWithFields } from './fields.ts';
import type { DocumentLayout, PlacedParagraph } from './page.ts';
import { fakeMeasurer, NO_GRID, numberLabel, para, run, SIZE_5 } from './test-fixtures.ts';

/** 一行 10 个字 */
const TEN = '一二三四五六七八九十';

/** 版心 10 字宽（2100）× 3 行高（819，行高 = 1.3 em） */
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

function opts(over: Partial<LayoutDocumentWithFieldsOptions> = {}): LayoutDocumentWithFieldsOptions {
  return { measurer: fakeMeasurer(), settings: DEFAULT_SETTINGS, ...over };
}

/**
 * 一个配对好的复杂域。真实来源是 `@uw/model` 的 `scanFields()`（界桩配对），
 * 这里直接造结果 —— 配对本身的正确性归 model 的 fields.test.ts 管。
 */
function field(instr: string, resultRuns: NodeId[]): FieldRegion {
  return {
    kind: 'complex',
    instruction: parseFieldInstruction(instr),
    instructionText: instr,
    depth: 0,
    resultRuns,
  };
}

/** 页上的第一块（都是段落，断言里不必每次重复这句判断） */
function firstPara(doc: DocumentLayout, page = 0): PlacedParagraph {
  const block = doc.pages[page]?.blocks[0];
  if (block === undefined || block.kind !== 'paragraph') throw new Error('这一页的第一块不是段落');
  return block;
}

/** 页上所有片段的文字，按顺序拼起来 —— 「第几页显示了什么」一眼可读 */
function pageText(doc: DocumentLayout, page: number): string {
  let out = '';
  for (const block of doc.pages[page]?.blocks ?? []) {
    if (block.kind !== 'paragraph') continue;
    for (const placed of block.lines) {
      for (const f of placed.line.fragments) out += f.text;
    }
  }
  return out;
}

describe('PAGE / NUMPAGES', () => {
  /** 三个域段落 + 垫行，排成「每页 1 个域 + 2 行垫料」 */
  function threePages(instr = 'PAGE'): {
    doc: DocumentLayout;
    runs: ResolvedRun[];
    fields: FieldRegion[];
  } {
    const runs = [run('9'), run('9'), run('9')];
    const pad = (): ResolvedBlock => para([run(TEN.repeat(2))]);
    const blocks: ResolvedBlock[] = [
      para([runs[0] as ResolvedRun]),
      pad(),
      para([runs[1] as ResolvedRun]),
      pad(),
      para([runs[2] as ResolvedRun]),
    ];
    const fields = runs.map((r) => field(instr, [r.id]));
    const res = layoutDocumentWithFields(body(blocks), fields, opts());
    return { doc: res.layout, runs, fields };
  }

  it('每页的 PAGE 域算出本页页码，盖掉文件里存着的旧值', () => {
    const { doc } = threePages();
    expect(doc.pages).toHaveLength(3);
    expect(pageText(doc, 0)).toBe(`1${TEN.repeat(2)}`);
    expect(pageText(doc, 1)).toBe(`2${TEN.repeat(2)}`);
    expect(pageText(doc, 2)).toBe('3');
  });

  it('NUMPAGES 数的是总页数', () => {
    const { doc } = threePages('NUMPAGES');
    expect(pageText(doc, 0).startsWith('3')).toBe(true);
    expect(pageText(doc, 2)).toBe('3');
  });

  it('SECTIONPAGES 只数本节的页', () => {
    const first = sect();
    const second = sect({ type: 'nextPage' });
    const r = run('9');
    const doc = layoutDocumentWithFields(
      {
        sections: [
          { id: 's0', props: first, blocks: [para([run(TEN.repeat(4))])] },
          { id: 's1', props: second, blocks: [para([r])] },
        ],
      },
      [field('SECTIONPAGES', [r.id])],
      opts(),
    );
    // 第一节 4 行 = 2 页，第二节 1 行 = 1 页
    expect(doc.layout.pages).toHaveLength(3);
    expect(pageText(doc.layout, 2)).toBe('1');
  });

  it('页码从 w:pgNumType w:start 起算时，PAGE 用的是显示页码而不是物理页序', () => {
    const r = run('9');
    const res = layoutDocumentWithFields(
      body([para([run(TEN.repeat(3))]), para([r])], sect({ pageNumStart: 7 })),
      [field('PAGE', [r.id])],
      opts(),
    );
    expect(pageText(res.layout, 1)).toBe('8');
  });

  it('结果区被切成几个 run 时，值落在第一个上、其余的旧值清掉', () => {
    const a = run('1');
    const b = run('2');
    const res = layoutDocumentWithFields(
      body([para([run(TEN.repeat(3))]), para([a, b])]),
      [field('PAGE', [a.id, b.id])],
      opts(),
    );
    // 旧值 "12" 整个被换成 "2"，而不是留下一个 "22"
    expect(pageText(res.layout, 1)).toBe('2');
  });

  it('域结果的片段带 field 标记（它不在 document.xml 里，反查不到 DocPosition）', () => {
    const r = run('9');
    const res = layoutDocumentWithFields(body([para([r])]), [field('PAGE', [r.id])], opts());
    const frag = firstPara(res.layout).lines[0]?.line.fragments[0];
    expect(frag?.text).toBe('1');
    expect(frag?.field).toBe(true);
  });
});

describe('数字格式', () => {
  function pageThree(instr: string, props = sect()): string {
    const r = run('9');
    const res = layoutDocumentWithFields(
      body([para([run(TEN.repeat(6))]), para([r])], props),
      [field(instr, [r.id])],
      opts(),
    );
    return pageText(res.layout, 2);
  }

  it('\\* ROMAN 与 \\* roman 按开关自己的大小写出大小写', () => {
    expect(pageThree('PAGE \\* ROMAN')).toBe('III');
    expect(pageThree('PAGE \\* roman')).toBe('iii');
  });

  it('\\* MERGEFORMAT 不是数字格式，退到十进制', () => {
    expect(pageThree('PAGE \\* MERGEFORMAT')).toBe('3');
  });

  it('一条指令里两个 \\*：取第一个认得出的那个', () => {
    expect(pageThree('PAGE \\* ROMAN \\* MERGEFORMAT')).toBe('III');
  });

  it('没写 \\* 时跟着本节的 w:pgNumType w:fmt', () => {
    expect(pageThree('PAGE', sect({ pageNumFormat: 'upperRoman' }))).toBe('III');
    // 域自己写了就以域为准
    expect(pageThree('PAGE \\* arabic', sect({ pageNumFormat: 'upperRoman' }))).toBe('3');
  });
});

describe('迭代与收敛', () => {
  it('没有可求值的域时只排一趟', () => {
    const r = run('2026-08-22');
    const res = layoutDocumentWithFields(body([para([r])]), [field('DATE', [r.id])], opts());
    expect(res.passes).toBe(1);
    expect(res.converged).toBe(true);
    expect(pageText(res.layout, 0)).toBe('2026-08-22');
  });

  it('一般情形两趟收敛（第一趟拿文件里的旧值，第二趟拿算出来的）', () => {
    const r = run('9');
    const res = layoutDocumentWithFields(body([para([r])]), [field('PAGE', [r.id])], opts());
    expect(res.passes).toBe(2);
    expect(res.converged).toBe(true);
  });

  it('域文字把内容顶出一页时，页码跟着变、再排一趟才自洽', () => {
    // 24 行整（8 页满），末行正好排满 10 个字；域的旧结果是空的
    const r = run('');
    const res = layoutDocumentWithFields(
      body([para([run(TEN.repeat(24)), r])]),
      [field('NUMPAGES', [r.id])],
      opts(),
    );
    // 第一趟 8 页 → 域算出 "8" → 末行放不下，多出第 25 行、多出第 9 页 → 域改算 "9" → 稳住
    expect(res.passes).toBe(3);
    expect(res.converged).toBe(true);
    expect(res.layout.pages).toHaveLength(9);
    expect(pageText(res.layout, 8).endsWith('9')).toBe(true);
  });

  it('撞上迭代上限时冻结在页数最多的那一趟，并记诊断', () => {
    const sink = createDiagnosticSink();
    const r = run('');
    const res = layoutDocumentWithFields(
      body([para([run(TEN.repeat(24)), r])]),
      [field('NUMPAGES', [r.id])],
      opts({ maxPasses: 1, diagnostics: sink }),
    );
    expect(res.converged).toBe(false);
    expect(sink.list().map((d) => d.code)).toContain('field-not-converged');
  });
});

describe('不求值的情形', () => {
  it('没有结果区的域什么都不显示，只记一条诊断', () => {
    const sink = createDiagnosticSink();
    const res = layoutDocumentWithFields(
      body([para([run(TEN)])]),
      [field('PAGE', [])],
      opts({ diagnostics: sink }),
    );
    expect(res.passes).toBe(1);
    expect(pageText(res.layout, 0)).toBe(TEN);
    expect(sink.list().map((d) => d.code)).toContain('field-no-result');
  });

  it('嵌套的可求值域按文档顺序先到先得，内层跳过并记诊断', () => {
    const sink = createDiagnosticSink();
    const r = run('9');
    const res = layoutDocumentWithFields(
      body([para([r])]),
      [field('PAGE', [r.id]), field('NUMPAGES', [r.id])],
      opts({ diagnostics: sink }),
    );
    expect(pageText(res.layout, 0)).toBe('1');
    expect(sink.list().map((d) => d.code)).toContain('field-nested-eval');
  });

  it('认不出的域原样显示文件里存着的旧结果', () => {
    const r = run('第 3 章');
    const res = layoutDocumentWithFields(body([para([r])]), [field('STYLEREF 1', [r.id])], opts());
    expect(pageText(res.layout, 0)).toBe('第 3 章');
  });
});

describe('PAGEREF（目录页码）', () => {
  /** 目录条目在第 1 页，标题（书签 `_Toc1` 所在段落）排在第 3 页 */
  function toc(instr: string, bookmarks?: ReadonlyMap<string, NodeId>, sink = createDiagnosticSink()) {
    const r = run('9');
    const heading = para([run('标题')]);
    const res = layoutDocumentWithFields(
      body([para([r]), para([run(TEN.repeat(5))]), heading]),
      [field(instr, [r.id])],
      opts({ bookmarks: bookmarks ?? new Map([['_Toc1', heading.id]]), diagnostics: sink }),
    );
    return { res, sink, entry: pageText(res.layout, 0).charAt(0) };
  }

  it('显示书签起点所在那一页的页码，不是域自己所在的页', () => {
    const { res, entry } = toc('PAGEREF _Toc1 \\h');
    expect(res.layout.pages).toHaveLength(3);
    expect(entry).toBe('3');
    expect(res.converged).toBe(true);
  });

  it('没写 \\* 时跟着**目标那一节**的页码格式', () => {
    const r = run('9');
    const heading = para([run('正文')]);
    const res = layoutDocumentWithFields(
      {
        sections: [
          { id: 's0', props: sect(), blocks: [para([r])] },
          {
            id: 's1',
            props: sect({ type: 'nextPage', pageNumFormat: 'lowerRoman', pageNumStart: 1 }),
            blocks: [heading],
          },
        ],
      },
      [field('PAGEREF _Toc1 \\h', [r.id])],
      opts({ bookmarks: new Map([['_Toc1', heading.id]]) }),
    );
    // 条目自己在第一节（阿拉伯数字），目标在第二节（罗马数字、从 1 起）
    expect(pageText(res.layout, 0)).toBe('i');
  });

  it('书签不存在时照旧显示存着的结果，并记 warn', () => {
    const { entry, sink } = toc('PAGEREF _Toc404 \\h');
    expect(entry).toBe('9');
    expect(sink.list().find((d) => d.code === 'field-bookmark-missing')?.severity).toBe('warn');
  });

  it('\\p（见上方 / 见下方）不求值', () => {
    const { entry, sink } = toc('PAGEREF _Toc1 \\p');
    expect(entry).toBe('9');
    expect(sink.list().map((d) => d.code)).toContain('field-pageref-relative');
  });

  it('没给书签表时一律不求值', () => {
    const r = run('9');
    const res = layoutDocumentWithFields(
      body([para([r]), para([run(TEN.repeat(5))]), para([run('标题')])]),
      [field('PAGEREF _Toc1', [r.id])],
      opts(),
    );
    expect(res.passes).toBe(1);
    expect(pageText(res.layout, 0).charAt(0)).toBe('9');
  });

  it('页码在两个解之间来回跳时认出环，不等撞上限就冻结在页数多的那一趟', () => {
    // 条目 = 8 个汉字 + 「x」+ 罗马数字页码，汉字与 x 之间有 1/4 字的中西文间距：
    // 页码「III」（1.5 字）共 10.25 字放不下、整个拉丁词换到下一行，「IV」（1 字）共 9.75 字一行放得下。
    // 条目一行时标题是第 9 行（第 3 页 → III → 条目变两行），两行时标题被挤到第 4 页
    // （→ IV → 条目变回一行）—— 两个解互相否定，没有不动点
    const sink = createDiagnosticSink();
    const r = run('IV');
    const heading = para([run('标题')]);
    const filler = Array.from({ length: 7 }, () => para([run(TEN)]));
    const res = layoutDocumentWithFields(
      body([para([run(`${TEN.slice(0, 8)}x`), r]), ...filler, heading]),
      [field('PAGEREF _Toc1 \\* ROMAN', [r.id])],
      opts({ bookmarks: new Map([['_Toc1', heading.id]]), diagnostics: sink }),
    );
    expect(res.converged).toBe(false);
    expect(res.passes).toBeLessThan(5);
    expect(res.layout.pages).toHaveLength(4);
    expect(res.values.get(r.id)).toBe('III');
    expect(sink.list().find((d) => d.code === 'field-not-converged')?.message).toContain('来回跳');
  });
});

describe('SEQ（题注编号）', () => {
  /** 一段「图 N」题注：`{ SEQ 图 … }` 的结果 run 存着旧值 9 */
  function caption(instr = 'SEQ 图 \\* ARABIC'): {
    block: ResolvedBlock;
    region: FieldRegion;
    run: ResolvedRun;
  } {
    const r = run('9');
    const block = para([run('图'), r]);
    return { block, region: field(instr, [r.id]), run: r };
  }
  const heading = (level: number): ResolvedBlock => para([run('标题')], { outlineLevel: level });

  it('按文档序数同名序列，不同名的各数各的；不进迭代', () => {
    const a = caption();
    const b = caption('SEQ 表');
    const c = caption();
    const res = layoutDocumentWithFields(
      body([a.block, b.block, c.block]),
      [a.region, b.region, c.region],
      opts(),
    );
    expect([a, b, c].map((x) => res.values.get(x.run.id))).toEqual(['1', '1', '2']);
    expect(pageText(res.layout, 0)).toBe('图1图1图2');
    expect(res.passes).toBe(1);
  });

  it('序列名不分大小写', () => {
    const a = caption('SEQ Figure');
    const b = caption('SEQ figure');
    const res = layoutDocumentWithFields(body([a.block, b.block]), [a.region, b.region], opts());
    expect(res.values.get(b.run.id)).toBe('2');
  });

  it('\\c 重复上一个号、\\r 重置、\\h 照数不显示', () => {
    const xs = [
      caption(),
      caption('SEQ 图 \\c'),
      caption('SEQ 图 \\h'),
      caption(),
      caption('SEQ 图 \\r 7'),
      caption(),
    ];
    const res = layoutDocumentWithFields(
      body(xs.map((x) => x.block)),
      xs.map((x) => x.region),
      opts(),
    );
    expect(xs.map((x) => res.values.get(x.run.id))).toEqual(['1', '1', '', '3', '7', '8']);
  });

  it('\\s N 遇到第 N 级（及更高）标题归零，只影响带它的那个域', () => {
    const [a, b, c, d] = [
      caption('SEQ 图 \\s 1'),
      caption('SEQ 图 \\s 1'),
      caption('SEQ 图 \\s 2'),
      caption('SEQ 图'),
    ];
    const xs = [a, b, c, d];
    const res = layoutDocumentWithFields(
      body([
        heading(0),
        a.block,
        b.block,
        heading(1),
        c.block, // 标题 2 → 从 1 数
        heading(0),
        d.block, // 不带 \s：接着数
      ]),
      xs.map((x) => x.region),
      opts(),
    );
    expect(xs.map((x) => res.values.get(x.run.id))).toEqual(['1', '2', '1', '2']);
  });

  it('\\s 2 遇到标题 1 也归零（更高级的标题开新一章）', () => {
    const a = caption('SEQ 图 \\s 2');
    const b = caption('SEQ 图 \\s 2');
    const res = layoutDocumentWithFields(body([a.block, heading(0), b.block]), [a.region, b.region], opts());
    expect(res.values.get(b.run.id)).toBe('1');
  });

  it('格式开关：\\* ROMAN / alphabetic', () => {
    const a = caption('SEQ 图 \\* ROMAN');
    const b = caption('SEQ 图 \\* alphabetic');
    const res = layoutDocumentWithFields(body([a.block, b.block]), [a.region, b.region], opts());
    expect([res.values.get(a.run.id), res.values.get(b.run.id)]).toEqual(['I', 'b']);
  });

  it('结果区拆成几个 run 时，只留第一个、其余清空', () => {
    const r1 = run('1');
    const r2 = run('0');
    const res = layoutDocumentWithFields(
      body([para([run('图'), r1, r2])]),
      [field('SEQ 图', [r1.id, r2.id])],
      opts(),
    );
    expect(pageText(res.layout, 0)).toBe('图1');
  });

  it('没有结果区的照样计数；带书签参数的不求值也不计数', () => {
    const sink = createDiagnosticSink();
    const empty = para([run('图')]);
    const bm = caption('SEQ 图 _Ref1');
    const last = caption();
    const noResult: FieldRegion = {
      ...field('SEQ 图', []),
      begin: { paragraphId: empty.id, runId: 'x', contentIndex: 0 },
    };
    const res = layoutDocumentWithFields(
      body([empty, bm.block, last.block]),
      [noResult, bm.region, last.region],
      opts({ diagnostics: sink }),
    );
    expect(res.values.get(last.run.id)).toBe('2');
    expect(res.values.has(bm.run.id)).toBe(false);
    expect(sink.list().some((d) => d.code === 'field-seq-bookmark')).toBe(true);
  });

  it('与 PAGE 一起迭代时，每一趟都带着 SEQ 的结果', () => {
    const a = caption();
    const p = run('9');
    const res = layoutDocumentWithFields(
      body([a.block, para([p])]),
      [a.region, field('PAGE', [p.id])],
      opts(),
    );
    expect(res.values.get(a.run.id)).toBe('1');
    expect(res.values.get(p.id)).toBe('1');
    expect(res.converged).toBe(true);
  });
});

describe('STYLEREF（章节号 / 页眉里的当前章）', () => {
  /** 标题 1 的样式 id 是 h1、w:name 是英文的 heading 1 —— 指令里写的却是中文界面名 */
  const names = new Map([
    ['h1', 'heading 1'],
    ['a', 'Normal'],
  ]);
  const h1 = (text: string, label?: string): ResolvedBlock =>
    para([run(text)], {
      styleId: 'h1',
      ...(label === undefined ? {} : { numbering: { numId: 1, level: 0, label: numberLabel(label) } }),
    });
  /** 一段「见 {STYLEREF …}」，结果 run 存着旧值「旧」 */
  function ref(instr: string): { block: ResolvedBlock; region: FieldRegion; run: ResolvedRun } {
    const r = run('旧');
    return { block: para([run('见'), r]), region: field(instr, [r.id]), run: r };
  }

  it('正文里往前找最近的那一段；前面没有才往后找', () => {
    const a = ref('STYLEREF "标题 1"');
    const b = ref('STYLEREF "标题 1"');
    const c = ref('STYLEREF "heading 1"');
    const res = layoutDocumentWithFields(
      body([a.block, h1('第一章'), b.block, h1('第二章'), c.block]),
      [a.region, b.region, c.region],
      opts({ styleNames: names }),
    );
    expect([a, b, c].map((x) => res.values.get(x.run.id))).toEqual(['第一章', '第一章', '第二章']);
    expect(res.passes).toBe(1);
  });

  it('数字 N 是「内建标题 N」的简写', () => {
    const a = ref('STYLEREF 1');
    const res = layoutDocumentWithFields(
      body([h1('第一章'), a.block]),
      [a.region],
      opts({ styleNames: names }),
    );
    expect(res.values.get(a.run.id)).toBe('第一章');
  });

  it('\\s / \\n \\t 只留编号里的数字，\\n 给整个编号', () => {
    const s1 = ref('STYLEREF 1 \\s');
    const n1 = ref('STYLEREF 1 \\n');
    const s2 = ref('STYLEREF 1 \\s');
    const t2 = ref('STYLEREF 1 \\n \\t');
    const res = layoutDocumentWithFields(
      body([h1('总则', '第1章'), s1.block, n1.block, h1('附则', '第一章'), s2.block, t2.block]),
      [s1.region, n1.region, s2.region, t2.region],
      opts({ styleNames: names }),
    );
    expect([s1, n1, s2, t2].map((x) => res.values.get(x.run.id))).toEqual(['1', '第1章', '一', '一']);
  });

  it('不求值的情形：样式不存在（记 warn）、要编号而没编号、\\p、没给样式表', () => {
    const sink = createDiagnosticSink();
    const missing = ref('STYLEREF "标题 7"');
    const noNumber = ref('STYLEREF 1 \\s');
    const relative = ref('STYLEREF 1 \\p');
    const res = layoutDocumentWithFields(
      body([h1('第一章'), missing.block, noNumber.block, relative.block]),
      [missing.region, noNumber.region, relative.region],
      opts({ styleNames: names, diagnostics: sink }),
    );
    for (const x of [missing, noNumber, relative]) expect(res.values.has(x.run.id)).toBe(false);
    expect(sink.list().some((d) => d.code === 'field-styleref-style-missing' && d.severity === 'warn')).toBe(
      true,
    );

    const bare = ref('STYLEREF 1');
    const res2 = layoutDocumentWithFields(body([h1('第一章'), bare.block]), [bare.region], opts());
    expect(res2.values.has(bare.run.id)).toBe(false);
  });

  describe('页眉里的', () => {
    const pad = (): ResolvedBlock => para([run(TEN)]);
    /** 第 1 页：第一章 + 两行；第 2 页：一行 + 第二章 + 一行；第 3 页：三行 */
    const blocks = (): ResolvedBlock[] => [
      h1('第一章'),
      pad(),
      pad(),
      pad(),
      h1('第二章'),
      pad(),
      pad(),
      pad(),
      pad(),
    ];
    function headerText(doc: DocumentLayout, page: number): string {
      let out = '';
      for (const block of doc.pages[page]?.header?.blocks ?? []) {
        if (block.kind !== 'paragraph') continue;
        for (const placed of block.lines) for (const f of placed.line.fragments) out += f.text;
      }
      return out;
    }
    function run3(instr: string) {
      const r = run('旧');
      const res = layoutDocumentWithFields(
        body(blocks(), sect({ headers: [{ type: 'default', relId: 'h' }] })),
        [field(instr, [r.id])],
        opts({ styleNames: names, headerFooters: { h: { resolved: [para([r])] } } }),
      );
      return res;
    }

    it('每页先在本页从上往下找，本页没有才往前找：一章从页中间开始时这一页已经是新的一章', () => {
      const res = run3('STYLEREF 1');
      expect(res.layout.pages).toHaveLength(3);
      expect([0, 1, 2].map((i) => headerText(res.layout, i))).toEqual(['第一章', '第二章', '第二章']);
      expect(res.converged).toBe(true);
    });

    it('\\l 在本页从下往上找（本页有两章时取后一章）', () => {
      const r = run('旧');
      const res = layoutDocumentWithFields(
        body(
          [h1('第一章'), h1('第二章'), pad(), pad(), pad(), pad()],
          sect({ headers: [{ type: 'default', relId: 'h' }] }),
        ),
        [field('STYLEREF 1 \\l', [r.id])],
        opts({ styleNames: names, headerFooters: { h: { resolved: [para([r])] } } }),
      );
      expect([0, 1].map((i) => headerText(res.layout, i))).toEqual(['第二章', '第二章']);
    });
  });
});

/** 合成度量器下 ASCII 是半角：这几个测试里「一个数字 = 半个汉字」的前提就靠它 */
it('前提自检：一行 10 个汉字正好排满版心', () => {
  const res = layoutDocumentWithFields(body([para([run(TEN)])]), [], opts());
  const line = firstPara(res.layout).lines[0]?.line;
  expect(line?.width).toBe(SIZE_5 * 10);
  expect(line?.fragments[0]?.text).toBe(TEN);
});
