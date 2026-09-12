/**
 * 查找 / 选择器 / 文档序。用手写的最小 XML，每个用例只打一个点。
 * 重点是「跨 run 能不能搜到」「隐藏 / 域 / 对象进不进串」「位置指得对不对」——
 * 搜到了却指错位置，比没搜到更糟（高亮画在别的字上）。
 */
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { CascadeContext } from './cascade.ts';
import type { Body, ResolvedBody } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { EMPTY_NUMBERING } from './numbering.ts';
import { buildRunOrder, compareDocPositions, rangeContains, rangeOfNode } from './order.ts';
import { parseBody } from './parse-body.ts';
import { parseSelector, queryNodes } from './query.ts';
import { resolveBody } from './resolve-body.ts';
import { findText } from './search.ts';
import { DEFAULT_SETTINGS } from './settings.ts';
import { parseStyles } from './styles.ts';
import { EMPTY_THEME } from './theme.ts';

const CTX: CascadeContext = {
  styles: parseStyles(
    parseXml(
      '<w:styles><w:style w:type="character" w:styleId="Secret"><w:rPr><w:vanish/></w:rPr></w:style></w:styles>',
    ),
    createDiagnosticSink(),
  ),
  theme: EMPTY_THEME,
  settings: DEFAULT_SETTINGS,
  numbering: EMPTY_NUMBERING,
};
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';

function parse(bodyXml: string): Body {
  const doc = parseXml(`<w:document><w:body>${bodyXml}${SECT}</w:body></w:document>`, 'document.xml');
  return parseBody(doc, createDiagnosticSink());
}
function resolved(bodyXml: string): ResolvedBody {
  return resolveBody(CTX, parse(bodyXml));
}
const p = (...runs: string[]) =>
  `<w:p>${runs.map((t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`).join('')}</w:p>`;

describe('findText', () => {
  it('跨 run 匹配，位置指到各自的 run', () => {
    const body = resolved(p('签', '发', '人：张三'));
    const hits = findText(body, '签发人');
    const runs = [...walkParagraphs(body)][0]?.runs ?? [];
    expect(hits).toEqual([
      {
        start: { nodeId: runs[0]?.id, contentIndex: 0, offset: 0 },
        end: { nodeId: runs[2]?.id, contentIndex: 0, offset: 1 },
      },
    ]);
  });

  it('不跨段落', () => {
    expect(findText(resolved(`${p('签发')}${p('人')}`), '签发人')).toEqual([]);
  });

  it('同一个 run 里挨着的两个 w:t 是两个片段，contentIndex 跟着走', () => {
    const body = resolved('<w:p><w:r><w:t>ab</w:t><w:t>cd</w:t></w:r></w:p>');
    const run = [...walkParagraphs(body)][0]?.runs[0];
    expect(findText(body, 'bc')).toEqual([
      {
        start: { nodeId: run?.id, contentIndex: 0, offset: 1 },
        end: { nodeId: run?.id, contentIndex: 1, offset: 1 },
      },
    ]);
  });

  it('字符串默认不分大小写，matchCase 才分；正则跟自己的 flags', () => {
    const body = resolved(p('Word word'));
    expect(findText(body, 'word')).toHaveLength(2);
    expect(findText(body, 'word', { matchCase: true })).toHaveLength(1);
    expect(findText(body, /word/)).toHaveLength(1);
    expect(findText(body, /word/i)).toHaveLength(2);
  });

  it('字符串里的正则元字符按字面量找', () => {
    expect(findText(resolved(p('1+1=2 (a)')), '1+1')).toHaveLength(1);
    expect(findText(resolved(p('1+1=2 (a)')), '(a)')).toHaveLength(1);
  });

  it('正则找全部，且不改调用方正则的 lastIndex；零长匹配跳过', () => {
    const re = /第\s*\d+\s*条/g;
    re.lastIndex = 3;
    const hits = findText(resolved(p('第1条、第 12 条')), re);
    expect(hits).toHaveLength(2);
    expect(re.lastIndex).toBe(3);
    expect(findText(resolved(p('aaa')), /x*/)).toEqual([]);
  });

  it('limit 截断', () => {
    expect(findText(resolved(p('aaaa')), 'a', { limit: 2 })).toHaveLength(2);
    expect(findText(resolved(p('aaaa')), 'a', { limit: 0 })).toEqual([]);
    expect(() => findText(resolved(p('a')), 'a', { limit: -1 })).toThrow(RangeError);
    expect(() => findText(resolved(p('a')), '')).toThrow(RangeError);
  });

  it('隐藏 run 不进串 —— 包括样式里写的 vanish', () => {
    const body = resolved(
      `<w:p><w:r><w:t>签发</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Secret"/></w:rPr><w:t>X</w:t></w:r><w:r><w:t>人</w:t></w:r></w:p>`,
    );
    expect(findText(body, '签发人')).toHaveLength(1);
    expect(findText(body, 'X')).toEqual([]);
  });

  it('域的界桩与指令不进串，结果区的旧值照搜；fieldValues 里的 run 整个跳过', () => {
    const body = resolved(
      `<w:p><w:r><w:t>第</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>3</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        <w:r><w:t>页</w:t></w:r></w:p>`,
    );
    expect(findText(body, '第3页')).toHaveLength(1);
    expect(findText(body, 'PAGE')).toEqual([]);
    const result = [...walkParagraphs(body)][0]?.runs[4];
    expect(findText(body, '第3页', { fieldValues: new Map([[result?.id ?? '', '4']]) })).toEqual([]);
  });

  it('制表位与换行各算一个字符，对象阻断匹配，软连字符不占位', () => {
    const body = resolved(
      `<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t><w:softHyphen/><w:t>d</w:t></w:r></w:p>`,
    );
    expect(findText(body, 'a\tb')).toHaveLength(1);
    expect(findText(body, 'b\nc')).toHaveLength(1);
    const cd = findText(body, 'cd');
    expect(cd).toHaveLength(1);
    // c 在第 4 片、d 在第 6 片（中间隔着软连字符），end 落在 d 那一片的末尾
    expect(cd[0]?.end).toEqual({ nodeId: cd[0]?.start.nodeId, contentIndex: 6, offset: 1 });

    const withImage = resolved(
      `<w:p><w:r><w:t>签发</w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="图片 1"/></wp:inline></w:drawing></w:r><w:r><w:t>人</w:t></w:r></w:p>`,
    );
    expect(findText(withImage, '签发人')).toEqual([]);
  });

  it('搜进表格单元格', () => {
    const body = resolved(`<w:tbl><w:tr><w:tc>${p('格内')}</w:tc></w:tr></w:tbl>${p('')}`);
    expect(findText(body, '格内')).toHaveLength(1);
  });
});

describe('queryNodes', () => {
  const DOC = `
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>标题</w:t></w:r></w:p>
    ${p('正文')}
    <w:tbl>
      <w:tr><w:tc>${p('表头')}</w:tc><w:tc>${p('表头2')}</w:tc></w:tr>
      <w:tr><w:tc>${p('a')}${p('b')}</w:tc><w:tc><w:tbl><w:tr><w:tc>${p('嵌套')}</w:tc></w:tr></w:tbl></w:tc></w:tr>
    </w:tbl>
    ${p('尾')}`;

  const text = (nodes: ReturnType<typeof queryNodes>) =>
    nodes.map((n) =>
      n.kind === 'paragraph' ? ((n.runs[0]?.content[0] as { text?: string })?.text ?? '') : n.kind,
    );

  it('类型 + 属性', () => {
    expect(text(queryNodes(parse(DOC), 'paragraph[styleId=Heading1]'))).toEqual(['标题']);
    expect(text(queryNodes(parse(DOC), 'paragraph[styleId="Heading1"]'))).toEqual(['标题']);
    expect(text(queryNodes(parse(DOC), 'paragraph[styleId]'))).toEqual(['标题']);
    expect(queryNodes(parse(DOC), 'table')).toHaveLength(2);
  });

  it('后代 vs 直接子；伪类数的是父列表里的位置', () => {
    const body = parse(DOC);
    expect(text(queryNodes(body, 'table > row:first-child cell'))).toEqual(['cell', 'cell', 'cell']);
    expect(text(queryNodes(body, 'table > row:first-child > cell > paragraph'))).toEqual([
      '表头',
      '表头2',
      '嵌套',
    ]);
    expect(text(queryNodes(body, 'row:nth-child(2) > cell:first-child > paragraph:last-child'))).toEqual([
      'b',
    ]);
    expect(text(queryNodes(body, 'cell paragraph'))).toEqual(['表头', '表头2', 'a', 'b', '嵌套']);
    // 顶层的 :first-child 数的是节里的块
    expect(text(queryNodes(body, 'paragraph:first-child'))).toEqual(['标题', '表头', '表头2', 'a', '嵌套']);
    expect(text(queryNodes(body, 'run'))).toHaveLength(8);
  });

  it('级联完的树也能查，属性按字符串比', () => {
    const body = resolved(DOC);
    expect(text(queryNodes(body, 'paragraph[styleId=Heading1]'))).toEqual(['标题']);
  });

  it('语法错误与不支持的类型抛错，且说清原因', () => {
    expect(() => parseSelector('')).toThrow(RangeError);
    expect(() => parseSelector('> paragraph')).toThrow(RangeError);
    expect(() => parseSelector('paragraph >')).toThrow(RangeError);
    expect(() => parseSelector('paragraph[styleId')).toThrow(RangeError);
    expect(() => parseSelector('paragraph:nth-child(0)')).toThrow(RangeError);
    expect(() => parseSelector('paragraph:hover')).toThrow(RangeError);
    expect(() => parseSelector('div')).toThrow(/不认识的类型/);
    expect(() => parseSelector('sdt[tag=applicant]')).toThrow(/内容控件/);
    expect(() => parseSelector('image')).toThrow(/不是节点/);
  });
});

describe('文档序与 rangeOf', () => {
  it('compare 按 run 序 → 片段 → 偏移；不在树里的 run 答 undefined', () => {
    const body = parse(`${p('ab', 'c')}${p('d')}`);
    const order = buildRunOrder(body);
    const [p1, p2] = [...walkParagraphs(body)];
    const r = (i: number, ci: number, offset: number, para = p1) => ({
      nodeId: para?.runs[i]?.id ?? '',
      contentIndex: ci,
      offset,
    });
    expect(compareDocPositions(order, r(0, 0, 1), r(0, 0, 2))).toBe(-1);
    expect(compareDocPositions(order, r(0, 0, 2), r(1, 0, 0))).toBe(-1);
    expect(compareDocPositions(order, r(0, 0, 0, p2), r(1, 0, 0))).toBe(1);
    expect(compareDocPositions(order, r(1, 0, 0), r(1, 0, 0))).toBe(0);
    expect(
      compareDocPositions(order, r(1, 0, 0), { nodeId: 'rId7:r0', contentIndex: 0, offset: 0 }),
    ).toBeUndefined();
  });

  it('rangeContains 是半开区间', () => {
    const body = parse(p('abc'));
    const order = buildRunOrder(body);
    const run = [...walkParagraphs(body)][0]?.runs[0];
    const at = (offset: number) => ({ nodeId: run?.id ?? '', contentIndex: 0, offset });
    const range = { start: at(0), end: at(2) };
    expect(rangeContains(order, range, at(0))).toBe(true);
    expect(rangeContains(order, range, at(1))).toBe(true);
    expect(rangeContains(order, range, at(2))).toBe(false);
    expect(rangeContains(order, range, { start: at(1), end: at(2) })).toBe(true);
    expect(rangeContains(order, range, { start: at(1), end: at(3) })).toBe(false);
  });

  it('rangeOfNode：段落 / 表格覆盖首 run 到末 run；空段落没有 range', () => {
    const body = parse(
      `${p('ab', 'cd')}<w:p/><w:tbl><w:tr><w:tc>${p('x')}</w:tc><w:tc>${p('yz')}</w:tc></w:tr></w:tbl>`,
    );
    const [para, empty] = [...walkParagraphs(body)];
    expect(rangeOfNode(para as NonNullable<typeof para>)).toEqual({
      start: { nodeId: para?.runs[0]?.id, contentIndex: 0, offset: 0 },
      end: { nodeId: para?.runs[1]?.id, contentIndex: 0, offset: 2 },
    });
    expect(rangeOfNode(empty as NonNullable<typeof empty>)).toBeUndefined();
    const table = queryNodes(body, 'table')[0];
    const cells = queryNodes(body, 'cell');
    expect(rangeOfNode(table as NonNullable<typeof table>)?.end).toEqual({
      nodeId: (cells[1] as { blocks: { runs: { id: string }[] }[] }).blocks[0]?.runs[0]?.id,
      contentIndex: 0,
      offset: 2,
    });
  });
});
