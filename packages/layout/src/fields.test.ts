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
import { fakeMeasurer, NO_GRID, para, run, SIZE_5 } from './test-fixtures.ts';

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

/** 合成度量器下 ASCII 是半角：这几个测试里「一个数字 = 半个汉字」的前提就靠它 */
it('前提自检：一行 10 个汉字正好排满版心', () => {
  const res = layoutDocumentWithFields(body([para([run(TEN)])]), [], opts());
  const line = firstPara(res.layout).lines[0]?.line;
  expect(line?.width).toBe(SIZE_5 * 10);
  expect(line?.fragments[0]?.text).toBe(TEN);
});
