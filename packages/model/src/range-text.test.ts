import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { walkParagraphs } from './nodes.ts';
import { EMPTY_NUMBERING } from './numbering.ts';
import { rangeOfNode } from './order.ts';
import { parseBody } from './parse-body.ts';
import { fragmentOfRange, textOfRange } from './range-text.ts';
import { resolveBody } from './resolve-body.ts';
import { DEFAULT_SETTINGS } from './settings.ts';
import { parseStyles } from './styles.ts';
import { EMPTY_THEME } from './theme.ts';

function required<T>(value: T | undefined): T {
  if (value === undefined) throw new Error('缺少测试节点');
  return value;
}

function setup(xml: string) {
  const sink = createDiagnosticSink();
  const body = resolveBody(
    {
      styles: parseStyles(
        parseXml(
          '<w:styles><w:style w:type="character" w:styleId="Secret"><w:rPr><w:vanish/></w:rPr></w:style></w:styles>',
        ),
        sink,
      ),
      theme: EMPTY_THEME,
      settings: DEFAULT_SETTINGS,
      numbering: EMPTY_NUMBERING,
    },
    parseBody(parseXml(`<w:document><w:body>${xml}</w:body></w:document>`), sink),
  );
  const paragraphs = [...walkParagraphs(body)];
  const ranges = paragraphs.map((p) => required(rangeOfNode(p)));
  const range = { start: required(ranges[0]).start, end: required(ranges.at(-1)).end };
  return { body, paragraphs, ranges, range };
}
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

describe('模型选区纯文本', () => {
  it('跨样式与片段精确截取 UTF-16 范围，反向选区一致', () => {
    const { body, range } = setup(
      '<w:p><w:r><w:t>甲e</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>\u0301😀</w:t><w:t>乙丙</w:t></w:r></w:p>',
    );
    range.start.offset = 1;
    range.end.offset = 1;
    expect(textOfRange(body, range)).toBe('e\u0301😀乙');
    expect(textOfRange(body, { start: range.end, end: range.start })).toBe('e\u0301😀乙');
  });

  it('保留空段与跨单元格段落，段尾到下一段开头只复制换行', () => {
    const { body, ranges, range } = setup(
      `${p('甲')}<w:p/>${p('乙')}<w:tbl><w:tr><w:tc>${p('丙')}</w:tc><w:tc>${p('丁')}</w:tc></w:tr></w:tbl>`,
    );
    expect(textOfRange(body, range)).toBe('甲\n\n乙\n丙\n丁');
    expect(textOfRange(body, { start: required(ranges[0]).end, end: required(ranges[1]).start })).toBe('\n');
    expect(textOfRange(body, required(ranges[1]))).toBe('');
  });

  it('跳过级联隐藏文字、域代码和软连字符，保留显式换行与制表位', () => {
    const { body, range } = setup(
      '<w:p><w:r><w:t>甲</w:t><w:tab/><w:br/><w:noBreakHyphen/><w:softHyphen/></w:r><w:r><w:rPr><w:rStyle w:val="Secret"/></w:rPr><w:t>隐藏</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>PAGE</w:instrText><w:fldChar w:fldCharType="separate"/><w:t>旧值</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>',
    );
    expect(textOfRange(body, range)).toBe('甲\t\n-旧值');
  });

  it('分页 / 分栏符在带格式片段里是 U+000C，纯文本里是换行', () => {
    const { body, range } = setup(
      '<w:p><w:r><w:t>甲</w:t><w:br w:type="page"/><w:t>乙</w:t><w:br w:type="column"/><w:br/></w:r></w:p>',
    );
    expect(fragmentOfRange(body, range).paragraphs[0]?.runs.map((r) => r.text)).toEqual(['甲\f乙\f\n']);
    expect(textOfRange(body, range)).toBe('甲\n乙\n\n');
  });

  it('域复制当前显示值，结果的其他 run 不重复输出', () => {
    const { body, paragraphs, range } = setup(
      '<w:p><w:r><w:t>第</w:t></w:r><w:r><w:t>99</w:t></w:r><w:r><w:t>9</w:t></w:r><w:r><w:t>页</w:t></w:r></w:p>',
    );
    const runs = required(paragraphs[0]).runs;
    const fields = new Map([
      [required(runs[1]).id, '12'],
      [required(runs[2]).id, ''],
    ]);
    expect(textOfRange(body, range, fields)).toBe('第12页');
    const after = { ...range, start: required(rangeOfNode(required(runs[1]))).end };
    expect(textOfRange(body, after, fields)).toBe('页');
  });

  it('符号按实际码点复制，对象占位阻断两侧文字，空域结果可复制', () => {
    const { body, paragraphs } = setup(p('甲'));
    const run = required(required(paragraphs[0]).runs[0]);
    run.content = [
      { kind: 'text', text: '甲' },
      { kind: 'symbol', font: 'Symbol', char: '😀多余' },
      { kind: 'object', objectKind: 'drawing', width: 100, height: 100 },
      { kind: 'text', text: '乙' },
    ];
    expect(textOfRange(body, required(rangeOfNode(run)))).toBe('甲😀\uFFFC乙');
    run.content = [];
    const next = { ...run, id: 'next', content: [{ kind: 'text' as const, text: '页' }] };
    required(paragraphs[0]).runs.push(next);
    const range = required(rangeOfNode(required(paragraphs[0])));
    expect(textOfRange(body, range, new Map([[run.id, '12']]))).toBe('12页');
  });

  it('拒绝外部位置与越界偏移，折叠选区不复制', () => {
    const { body, range } = setup(p('甲'));
    expect(textOfRange(body, { start: range.start, end: range.start })).toBe('');
    for (const start of [
      { ...range.start, nodeId: 'missing' },
      { ...range.start, offset: 2 },
      { ...range.start, offset: -1 },
      { ...range.start, contentIndex: 1 },
    ])
      expect(() => textOfRange(body, { start, end: range.end })).toThrow(RangeError);
  });

  it('带格式片段：级联后的格式与段落对齐，同一 run 的内容片拼回一段，与纯文本同一套取舍', () => {
    const { body, range } = setup(
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>甲</w:t><w:tab/><w:t>乙</w:t></w:r>' +
        '<w:r><w:rPr><w:rStyle w:val="Secret"/></w:rPr><w:t>藏</w:t></w:r><w:r><w:t>丙</w:t></w:r></w:p><w:p/>',
    );
    const fragment = fragmentOfRange(body, range);
    expect(fragment.paragraphs.map((p) => p.justification)).toEqual(['center', 'left']);
    expect(fragment.paragraphs[0]?.runs.map((r) => [r.text, r.props.bold])).toEqual([
      ['甲\t乙', true],
      ['丙', false],
    ]);
    expect(fragment.paragraphs[1]?.runs).toEqual([]);
    expect(structuredClone(fragment)).toEqual(fragment);
    expect(fragmentOfRange(body, { start: range.start, end: range.start })).toEqual({ paragraphs: [] });
  });
});
