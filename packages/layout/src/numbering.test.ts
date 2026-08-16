/**
 * 列表编号的行内几何。
 *
 * 编号是一段**没有 run 的文字**：排在首行最前面、不能被断开、不参与两端对齐的拉伸、
 * 不能被当成可编辑内容。这四条各有一个用例。
 *
 * 版心宽度一律 `SIZE_5 * 10`（一行 10 个汉字），期望值数着字写。
 */
import { DEFAULT_SETTINGS } from '@uw/model';
import { describe, expect, it } from 'vitest';
import type { LayoutParagraphOptions } from './paragraph.ts';
import { layoutParagraph } from './paragraph.ts';
import { fakeMeasurer, NO_GRID, numberLabel, para, paraProps, run, SIZE_5 } from './test-fixtures.ts';
import type { LineLayout } from './types.ts';

const WIDTH = SIZE_5 * 10;

function opts(over: Partial<LayoutParagraphOptions> = {}): LayoutParagraphOptions {
  return {
    measurer: fakeMeasurer(),
    contentWidth: WIDTH,
    settings: DEFAULT_SETTINGS,
    docGrid: NO_GRID,
    ...over,
  };
}

/** 悬挂缩进 3 个字：编号从版心左边起，正文从 3 个字宽处起 —— 列表最常见的排法 */
function hanging(chars: number) {
  return { ...paraProps().indent, left: SIZE_5 * chars, hanging: SIZE_5 * chars };
}

const numberingFragments = (line: LineLayout | undefined) =>
  (line?.fragments ?? []).filter((f) => f.numbering === true);
const bodyFragments = (line: LineLayout | undefined) =>
  (line?.fragments ?? []).filter((f) => f.numbering !== true);

describe('编号的位置', () => {
  it('编号排在首行最前面，正文被制表位顶到左缩进上', () => {
    const p = para([run('正文')], {
      indent: hanging(3),
      numbering: { numId: 1, level: 0, label: numberLabel('一、') },
    });
    const line = layoutParagraph(p, opts()).lines[0];

    // 首行左边缘 = left - hanging = 0，编号就从那儿开始
    expect(numberingFragments(line)[0]).toMatchObject({ text: '一、', x: 0 });
    // 编号占 2 个字宽（420），制表位把正文送到左缩进 630 —— 而不是默认制表位的 720
    expect(bodyFragments(line)[0]?.x).toBe(SIZE_5 * 3);
  });

  it('编号宽过悬挂缩进时，隐含停靠点失效，退回普通制表位规则', () => {
    const p = para([run('正文')], {
      // 悬挂缩进只有 1 个字，编号却有 4 个字宽
      indent: hanging(1),
      numbering: { numId: 1, level: 0, label: numberLabel('一二三四') },
    });
    const line = layoutParagraph(p, opts()).lines[0];
    // 编号右边缘 840 已越过 210，落到 defaultTabStop（720）的下一个整数倍
    expect(bodyFragments(line)[0]?.x).toBe(DEFAULT_SETTINGS.defaultTabStop * 2);
  });

  it('w:suff=space / nothing：不走制表位，正文紧跟编号', () => {
    const base = { indent: hanging(3) };
    const space = para([run('正文')], {
      ...base,
      numbering: { numId: 1, level: 0, label: numberLabel('一、', { suffix: 'space' }) },
    });
    const nothing = para([run('正文')], {
      ...base,
      numbering: { numId: 1, level: 0, label: numberLabel('一、', { suffix: 'nothing' }) },
    });
    // 空格是半角（合成度量器里 ASCII 半角）
    expect(bodyFragments(layoutParagraph(space, opts()).lines[0])[0]?.x).toBe(SIZE_5 * 2 + SIZE_5 / 2);
    expect(bodyFragments(layoutParagraph(nothing, opts()).lines[0])[0]?.x).toBe(SIZE_5 * 2);
  });

  it('编号文字为空（numFmt=none）时分隔符照留 —— 正文仍落在左缩进上', () => {
    const p = para([run('正文')], {
      indent: hanging(3),
      numbering: { numId: 1, level: 0, label: numberLabel('') },
    });
    const line = layoutParagraph(p, opts()).lines[0];
    expect(numberingFragments(line)).toEqual([]);
    expect(bodyFragments(line)[0]?.x).toBe(SIZE_5 * 3);
  });

  it('没有 label 的段落一切照旧 —— 编号是可选的一段前缀，不是必经之路', () => {
    const p = para([run('正文')], { indent: hanging(3) });
    const line = layoutParagraph(p, opts()).lines[0];
    expect(numberingFragments(line)).toEqual([]);
    expect(line?.fragments[0]?.x).toBe(0); // 首行仍从 left - hanging 起
  });
});

describe('编号与断行', () => {
  it('编号只待在首行，且内部不被拆开', () => {
    const p = para([run('一二三四五六七八九十一二三四五')], {
      numbering: { numId: 1, level: 0, label: numberLabel('壹贰叁') },
    });
    const out = layoutParagraph(p, opts());
    // 编号 3 字 + 制表位到 720 + 正文，第一行装不下的正文换行；编号全在第 0 行
    expect(numberingFragments(out.lines[0])).toHaveLength(1);
    for (const line of out.lines.slice(1)) expect(numberingFragments(line)).toEqual([]);
    expect(out.lines[0]?.start).toBe(0);
  });

  it('编号的字号参与行高 —— 大一号的项目符号会把首行撑高', () => {
    const big = para([run('正文')], {
      numbering: {
        numId: 1,
        level: 0,
        label: numberLabel('•', { runProps: { ...paraProps().markRunProps, size: SIZE_5 * 2 } }),
      },
    });
    const plain = para([run('正文')]);
    const h = (p: typeof big) => layoutParagraph(p, opts()).lines[0]?.height ?? 0;
    expect(h(big)).toBeGreaterThan(h(plain));
  });
});

describe('编号不是正文', () => {
  it('两端对齐把多余宽度摊给正文，编号与它的制表位纹丝不动', () => {
    const p = para([run('中文English中文English中')], {
      justification: 'both',
      numbering: { numId: 1, level: 0, label: numberLabel('一、') },
    });
    const line = layoutParagraph(p, opts()).lines[0];
    const label = numberingFragments(line)[0];
    // 两个编号字仍在 0 与 210，一点没被拉开
    expect(label?.glyphX).toEqual([0, SIZE_5]);
  });

  it('片段带 numbering 标记、runId 指向段落而非某个 run —— 可选文本层要靠它把编号排除', () => {
    const p = para([run('正文')], { numbering: { numId: 1, level: 0, label: numberLabel('一、') } });
    const label = numberingFragments(layoutParagraph(p, opts()).lines[0])[0];
    expect(label?.numbering).toBe(true);
    expect(label?.runId).toBe(`${p.id}#num`);
    expect(p.runs.some((r) => r.id === label?.runId)).toBe(false);
  });

  it('「一个字符」的尺子取正文的字号，不取编号的 —— 首行缩进不该被项目符号带偏', () => {
    const p = para([run('正文')], {
      indent: { ...paraProps().indent, firstLineChars: 200 },
      numbering: {
        numId: 1,
        level: 0,
        label: numberLabel('•', {
          suffix: 'nothing',
          runProps: { ...paraProps().markRunProps, size: SIZE_5 * 4 },
        }),
      },
    });
    // 缩进 = 2 × 正文字号，而不是 2 × 编号字号
    expect(layoutParagraph(p, opts()).lines[0]?.x).toBe(SIZE_5 * 2);
  });
});
