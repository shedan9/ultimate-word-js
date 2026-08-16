/**
 * 编号的**消费**：计数器（文档顺序）+ 编号那一层接进级联。
 *
 * 与 parts.test.ts 分开：那边测的是 `numbering.xml` 的解析与解引用（纯查表），
 * 这边测的是「第 3 段该显示几」以及「编号的缩进有没有铺到段落上」——
 * 前者无状态，后者是整份文档跑一遍才有的结果，混在一个文件里会让人以为解析也有状态。
 *
 * `gongwen-01.docx` 没有 numbering.xml，所以这里全是手写样本；真实公文的编号
 * 要等语料库里进一份带多级列表的文档（DEVELOPMENT-PLAN §7 第 9 步）。
 */
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { CascadeContext } from './cascade.ts';
import { resolveParaProps } from './cascade.ts';
import type { Body, ResolvedParagraph } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { parseNumbering } from './numbering.ts';
import { createNumberingCounters } from './numbering-counter.ts';
import { parseBody } from './parse-body.ts';
import { parseParaProps } from './parse-props.ts';
import { resolveBody } from './resolve-body.ts';
import { DEFAULT_SETTINGS } from './settings.ts';
import { parseStyles } from './styles.ts';
import { EMPTY_THEME } from './theme.ts';

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function numberingFrom(xml: string) {
  return parseNumbering(parseXml(`<w:numbering ${W_NS}>${xml}</w:numbering>`), createDiagnosticSink());
}

function ctxFrom(numberingXml: string, stylesXml = ''): CascadeContext {
  const sink = createDiagnosticSink();
  return {
    styles: parseStyles(parseXml(`<w:styles ${W_NS}>${stylesXml}</w:styles>`), sink),
    theme: EMPTY_THEME,
    settings: DEFAULT_SETTINGS,
    numbering: numberingFrom(numberingXml),
  };
}

/** 一份三级列表定义：ilvl 0 用中文数字，1 用十进制，2 用小写字母 */
const THREE_LEVELS = `
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%1、"/>
      <w:pPr><w:ind w:left="640" w:hanging="640"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/><w:sz w:val="32"/></w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
    <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%3)"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>`;

describe('计数器', () => {
  it('同一级连着数；换到浅一级时深的级归零', () => {
    const c = createNumberingCounters(numberingFrom(THREE_LEVELS));
    expect(c.advance(1, 0)?.text).toBe('一、');
    // %1 用的是**第 0 级自己的** numFmt（中文），%2 用第 1 级的（十进制）——
    // 每级各按各的格式出，这是 lvlText 最容易想当然的一处
    expect(c.advance(1, 1)?.text).toBe('一.1');
    expect(c.advance(1, 1)?.text).toBe('一.2');
    expect(c.advance(1, 2)?.text).toBe('a)');
    expect(c.advance(1, 0)?.text).toBe('二、'); // 第 0 级 +1
    expect(c.advance(1, 1)?.text).toBe('二.1'); // 第 1 级归零后重来
    expect(c.advance(1, 2)?.text).toBe('a)');
  });

  it('计数按 numId 分家 —— 同一个 abstractNum 的两个实例各数各的', () => {
    const c = createNumberingCounters(numberingFrom(THREE_LEVELS));
    expect(c.advance(1, 0)?.text).toBe('一、');
    expect(c.advance(1, 0)?.text).toBe('二、');
    // numId=2 指向同一份定义，但它是另一个实例，从头开始
    expect(c.advance(2, 0)?.text).toBe('一、');
    expect(c.advance(1, 0)?.text).toBe('三、');
  });

  it('跳级引用取上一级的 start，且**不**把那一级用掉', () => {
    const c = createNumberingCounters(numberingFrom(THREE_LEVELS));
    // 文档一上来就是第二级：第 0 级按 start 显示成「一」，但它还没被用过
    expect(c.advance(1, 1)?.text).toBe('一.1');
    expect(c.advance(1, 1)?.text).toBe('一.2');
    expect(c.advance(1, 0)?.text).toBe('一、'); // 仍然是「一」，不是「二」
  });

  it('w:startOverride 让实例从别处起头', () => {
    const n = numberingFrom(`
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
      <w:num w:numId="2"><w:abstractNumId w:val="0"/>
        <w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>
      </w:num>`);
    const c = createNumberingCounters(n);
    expect(c.advance(1, 0)?.text).toBe('1.');
    expect(c.advance(2, 0)?.text).toBe('5.');
    expect(c.advance(2, 0)?.text).toBe('6.');
  });

  it('w:lvlRestart=0 从不归零：第 1 级跨过第 0 级的变化继续数', () => {
    const n = numberingFrom(`
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/></w:lvl>
        <w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/><w:lvlText w:val="%2"/><w:lvlRestart w:val="0"/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`);
    const c = createNumberingCounters(n);
    c.advance(1, 0);
    expect(c.advance(1, 1)?.text).toBe('1');
    c.advance(1, 0);
    expect(c.advance(1, 1)?.text).toBe('2');
  });

  it('w:isLgl 把所有层级压成阿拉伯数字，本级的 numFmt 也不例外', () => {
    const n = numberingFrom(`
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%1、"/></w:lvl>
        <w:lvl w:ilvl="1"><w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%1.%2"/><w:isLgl/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`);
    const c = createNumberingCounters(n);
    expect(c.advance(1, 0)?.text).toBe('一、');
    expect(c.advance(1, 1)?.text).toBe('1.1');
  });

  it('numId=0（取消编号）与指向不存在的定义：不编号，也不推进计数器', () => {
    const c = createNumberingCounters(numberingFrom(THREE_LEVELS));
    expect(c.advance(1, 0)?.text).toBe('一、');
    expect(c.advance(0, 0)).toBeUndefined();
    expect(c.advance(99, 0)).toBeUndefined();
    expect(c.advance(1, 0)?.text).toBe('二、');
  });

  it('bullet 的 lvlText 是字面字符，不当模板展开', () => {
    const n = numberingFrom(`
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`);
    const c = createNumberingCounters(n);
    expect(c.advance(1, 0)?.text).toBe('');
    expect(c.advance(1, 0)?.value).toBe(2); // 符号不变，但计数照走 —— 交叉引用要这个数
  });
});

describe('编号接进级联', () => {
  const ctx = ctxFrom(THREE_LEVELS);
  const numPr = (numId: number, ilvl: number) =>
    parseParaProps(
      parseXml(
        `<w:pPr ${W_NS}><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>`,
      ).root,
    );

  it('编号级的 pPr 铺到段落上 —— 悬挂缩进来自 numbering.xml，不是段落自己写的', () => {
    const p = resolveParaProps(ctx, numPr(1, 0));
    expect(p.indent.left).toBe(640);
    expect(p.indent.hanging).toBe(640);
  });

  it('直接格式压过编号级 —— 编号层在样式之后、直接格式之前', () => {
    const direct = parseParaProps(
      parseXml(
        `<w:pPr ${W_NS}><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="1000"/></w:pPr>`,
      ).root,
    );
    const p = resolveParaProps(ctx, direct);
    expect(p.indent.left).toBe(1000);
    expect(p.indent.hanging).toBe(640); // 编号级给的那半仍在（逐属性合并）
  });

  it('编号级的 pPr 压过段落样式', () => {
    const styled = ctxFrom(
      THREE_LEVELS,
      `<w:style w:type="paragraph" w:styleId="L"><w:name w:val="L"/><w:pPr><w:ind w:left="99"/></w:pPr></w:style>`,
    );
    const direct = parseParaProps(
      parseXml(
        `<w:pPr ${W_NS}><w:pStyle w:val="L"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>`,
      ).root,
    );
    expect(resolveParaProps(styled, direct).indent.left).toBe(640);
  });

  it('样式里声明的 numPr 一样算数（公文里编号常挂在样式上）', () => {
    const styled = ctxFrom(
      THREE_LEVELS,
      `<w:style w:type="paragraph" w:styleId="L"><w:name w:val="L"/>
        <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>`,
    );
    const p = resolveParaProps(
      styled,
      parseParaProps(parseXml(`<w:pPr ${W_NS}><w:pStyle w:val="L"/></w:pPr>`).root),
    );
    expect(p.numbering).toMatchObject({ numId: 1, level: 0 });
    expect(p.indent.hanging).toBe(640);
  });

  it('不传计数器时有缩进但没有 label —— 「第几」要靠前文才知道', () => {
    const p = resolveParaProps(ctx, numPr(1, 0));
    expect(p.indent.left).toBe(640);
    expect(p.numbering.label).toBeUndefined();
  });

  it('编号的 rPr 只作用于编号文字，不碰正文', () => {
    const counters = createNumberingCounters(ctx.numbering);
    const p = resolveParaProps(ctx, numPr(1, 0), counters);
    expect(p.numbering.label?.text).toBe('一、');
    expect(p.numbering.label?.runProps.fonts.ascii).toBe('Symbol');
    expect(p.numbering.label?.runProps.size).toBe(320); // 16pt
    // 段落标记（也就是正文那条字符级联）没有被 Symbol 污染
    expect(p.markRunProps.fonts.ascii).toBe('');
    expect(p.markRunProps.size).not.toBe(320);
  });

  it('段落标记的 rPr 压过编号级的 rPr —— 在 Word 里选中 ¶ 调字号，编号跟着变', () => {
    const direct = parseParaProps(
      parseXml(
        `<w:pPr ${W_NS}><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
         <w:rPr><w:sz w:val="48"/></w:rPr></w:pPr>`,
      ).root,
    );
    const label = resolveParaProps(ctx, direct, createNumberingCounters(ctx.numbering)).numbering.label;
    expect(label?.runProps.size).toBe(480);
    expect(label?.runProps.fonts.ascii).toBe('Symbol'); // 字体仍来自编号级
  });

  it('w:suff 与 w:lvlJc 带到 label 上；两者的缺省是 tab / left', () => {
    const n = ctxFrom(`
      <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
          <w:suff w:val="space"/><w:lvlJc w:val="right"/></w:lvl>
      </w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`);
    const label = resolveParaProps(n, numPr(1, 0), createNumberingCounters(n.numbering)).numbering.label;
    expect(label).toMatchObject({ suffix: 'space', justification: 'right', value: 1 });
    expect(
      resolveParaProps(ctx, numPr(1, 0), createNumberingCounters(ctx.numbering)).numbering.label,
    ).toMatchObject({ suffix: 'tab', justification: 'left' });
  });
});

describe('resolveBody 按文档顺序推进', () => {
  function bodyOf(xml: string): Body {
    return parseBody(
      parseXml(`<w:document ${W_NS}><w:body>${xml}<w:sectPr/></w:body></w:document>`, 'document.xml'),
      createDiagnosticSink(),
    );
  }
  const p = (numId: number, ilvl: number, text: string) =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>
      <w:r><w:t>${text}</w:t></w:r></w:p>`;

  const labels = (body: Body, ctx: CascadeContext): (string | undefined)[] =>
    [...walkParagraphs(resolveBody(ctx, body))].map(
      (para) => (para as ResolvedParagraph).props.numbering.label?.text,
    );

  it('整份文档一趟走完，编号连着数（普通段落不打断）', () => {
    const ctx = ctxFrom(THREE_LEVELS);
    const body = bodyOf(
      p(1, 0, '一节') +
        p(1, 1, '细目') +
        '<w:p><w:r><w:t>没编号的段</w:t></w:r></w:p>' +
        p(1, 1, '细目') +
        p(1, 0, '二节'),
    );
    expect(labels(body, ctx)).toEqual(['一、', '一.1', undefined, '一.2', '二、']);
  });

  it('表格单元格里的段落算在同一条计数里 —— 版记表格中的列表不另起一套', () => {
    const ctx = ctxFrom(THREE_LEVELS);
    const body = bodyOf(
      `${p(1, 0, '正文一')}<w:tbl><w:tr><w:tc>${p(1, 0, '表内')}</w:tc></w:tr></w:tbl>${p(1, 0, '正文二')}`,
    );
    expect(labels(body, ctx)).toEqual(['一、', '二、', '三、']);
  });

  it('同一份文档解析两次结果相同 —— 计数器不跨调用残留', () => {
    const ctx = ctxFrom(THREE_LEVELS);
    const body = bodyOf(p(1, 0, 'a') + p(1, 0, 'b'));
    expect(labels(body, ctx)).toEqual(labels(body, ctx));
  });
});
