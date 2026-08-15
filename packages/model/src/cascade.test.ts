/**
 * 级联的回归测试。
 *
 * 前半用手写的极小样式表打单点（顺序、逐属性合并、成环），
 * 后半直接跑真实公文 `gongwen-01.docx` —— 手写样本会把 Word 的实际习惯全绕开，
 * 比如「主题字体的 a:ea 是空的」这种坑，只有真文件才暴露得出来。
 */
import { createDiagnosticSink, twipsToPt } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { CascadeContext } from './cascade.ts';
import { resolveParaProps, resolveRunProps } from './cascade.ts';
import { parseParaProps, parseRunProps } from './parse-props.ts';
import { DEFAULT_SETTINGS } from './settings.ts';
import { parseStyles } from './styles.ts';
import { EMPTY_THEME } from './theme.ts';

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** 用一段 styles.xml 片段搭出级联上下文 */
function ctxFrom(stylesXml: string): CascadeContext {
  const sink = createDiagnosticSink();
  const doc = parseXml(`<w:styles ${W_NS}>${stylesXml}</w:styles>`);
  return { styles: parseStyles(doc, sink), theme: EMPTY_THEME, settings: DEFAULT_SETTINGS };
}

const pPr = (xml: string) => parseParaProps(parseXml(`<w:pPr ${W_NS}>${xml}</w:pPr>`).root);
const rPr = (xml: string) => parseRunProps(parseXml(`<w:rPr ${W_NS}>${xml}</w:rPr>`).root);

describe('级联顺序', () => {
  const ctx = ctxFrom(`
    <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/><w:b/></w:rPr></w:rPrDefault></w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
      <w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="H1">
      <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
      <w:pPr><w:jc w:val="center"/><w:keepNext/></w:pPr>
      <w:rPr><w:sz w:val="48"/></w:rPr>
    </w:style>
    <w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:sz w:val="10"/></w:rPr></w:style>
  `);

  it('后代样式覆盖祖先 —— basedOn 链的方向不能反', () => {
    const p = resolveParaProps(ctx, pPr('<w:pStyle w:val="H1"/>'));
    expect(p.justification).toBe('center'); // H1 覆盖了 Normal 的 both
    expect(p.keepNext).toBe(true);
  });

  it('没写 pStyle 时套用 w:default="1" 的段落样式', () => {
    const p = resolveParaProps(ctx, pPr(''));
    expect(p.styleId).toBe('Normal');
    expect(p.justification).toBe('both');
  });

  it('直接格式压过样式', () => {
    const p = resolveParaProps(ctx, pPr('<w:pStyle w:val="H1"/><w:jc w:val="right"/>'));
    expect(p.justification).toBe('right');
  });

  it('docDefaults 是最底层，没人覆盖时生效', () => {
    const r = resolveRunProps(ctx, undefined, undefined);
    expect(twipsToPt(r.size)).toBe(10.5); // w:sz=21 半磅
    expect(r.bold).toBe(true);
  });

  it('段落样式携带的字符属性也参与字符级联', () => {
    const r = resolveRunProps(ctx, pPr('<w:pStyle w:val="H1"/>'), undefined);
    expect(twipsToPt(r.size)).toBe(24); // H1 的 w:sz=48 半磅
  });

  it('字符样式排在段落样式之后', () => {
    const r = resolveRunProps(ctx, pPr('<w:pStyle w:val="H1"/>'), rPr('<w:rStyle w:val="Strong"/>'));
    expect(twipsToPt(r.size)).toBe(5); // 字符样式的 w:sz=10 赢
  });

  it('run 的直接格式压过一切', () => {
    const r = resolveRunProps(
      ctx,
      pPr('<w:pStyle w:val="H1"/>'),
      rPr('<w:rStyle w:val="Strong"/><w:sz w:val="30"/>'),
    );
    expect(twipsToPt(r.size)).toBe(15);
  });

  it('w:b w:val="0" 是「明确关掉」，能盖住 docDefaults 的 <w:b/>', () => {
    expect(resolveRunProps(ctx, undefined, rPr('<w:b w:val="0"/>')).bold).toBe(false);
    // 而「没写」是继承，不是关掉
    expect(resolveRunProps(ctx, undefined, rPr('<w:i/>')).bold).toBe(true);
  });
});

describe('逐属性合并', () => {
  const ctx = ctxFrom(`
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
      <w:name w:val="Normal"/>
      <w:pPr><w:ind w:left="420" w:right="200"/><w:spacing w:before="100" w:line="360" w:lineRule="auto"/></w:pPr>
    </w:style>
  `);

  it('w:ind 的属性各自继承，不整块替换', () => {
    const p = resolveParaProps(ctx, pPr('<w:ind w:firstLineChars="200" w:firstLine="640"/>'));
    expect(p.indent.left).toBe(420); // 样式里的没被吃掉
    expect(p.indent.right).toBe(200);
    expect(p.indent.firstLineChars).toBe(200);
  });

  it('w:spacing 同理，且 line/lineRule 不换算单位', () => {
    const p = resolveParaProps(ctx, pPr('<w:spacing w:after="480"/>'));
    expect(p.spacing.before).toBe(100);
    expect(p.spacing.after).toBe(480);
    expect(p.spacing.line).toBe(360); // 1/240 行，不是 twips
    expect(p.spacing.lineRule).toBe('auto');
  });

  it('rFonts 的四个桶各自继承', () => {
    const c = ctxFrom(`
      <w:docDefaults><w:rPrDefault><w:rPr>
        <w:rFonts w:ascii="Calibri" w:eastAsia="宋体"/>
      </w:rPr></w:rPrDefault></w:docDefaults>
    `);
    const r = resolveRunProps(c, undefined, rPr('<w:rFonts w:ascii="Times New Roman" w:hint="eastAsia"/>'));
    expect(r.fonts.ascii).toBe('Times New Roman');
    expect(r.fonts.eastAsia).toBe('宋体'); // 没被覆盖
    expect(r.fonts.hint).toBe('eastAsia');
  });
});

describe('默认值', () => {
  const ctx = ctxFrom('');

  it('Word 默认开着的那几项确实是开的', () => {
    const p = resolveParaProps(ctx, undefined);
    expect(p.widowControl).toBe(true);
    expect(p.autoSpaceDE).toBe(true);
    expect(p.autoSpaceDN).toBe(true);
    expect(p.overflowPunct).toBe(true);
    expect(p.snapToGrid).toBe(true);
  });

  it('szCs 没写时跟随 sz，不是退回 10pt', () => {
    const r = resolveRunProps(ctx, undefined, rPr('<w:sz w:val="44"/>'));
    expect(r.sizeCs).toBe(r.size);
    expect(twipsToPt(r.size)).toBe(22);
  });

  it('Resolved 属性可结构化克隆（原则 1.1）', () => {
    const p = resolveParaProps(ctx, pPr('<w:jc w:val="center"/>'));
    expect(structuredClone(p)).toEqual(p);
  });
});

describe('basedOn 成环', () => {
  it('记诊断并截断，不抛也不死循环', () => {
    const sink = createDiagnosticSink();
    const doc = parseXml(`<w:styles ${W_NS}>
      <w:style w:type="paragraph" w:styleId="A"><w:name w:val="A"/><w:basedOn w:val="B"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="B"><w:name w:val="B"/><w:basedOn w:val="A"/></w:style>
    </w:styles>`);
    const ctx: CascadeContext = {
      styles: parseStyles(doc, sink),
      theme: EMPTY_THEME,
      settings: DEFAULT_SETTINGS,
    };

    const p = resolveParaProps(ctx, pPr('<w:pStyle w:val="A"/>'));
    expect(p.justification).toBe('center');
    expect(sink.list().map((d) => d.code)).toContain('style-cycle');
  });

  it('basedOn 指向不存在的样式：记诊断，不影响其余属性', () => {
    const sink = createDiagnosticSink();
    const doc = parseXml(`<w:styles ${W_NS}>
      <w:style w:type="paragraph" w:styleId="A"><w:name w:val="A"/><w:basedOn w:val="没这个"/><w:pPr><w:jc w:val="right"/></w:pPr></w:style>
    </w:styles>`);
    const ctx: CascadeContext = {
      styles: parseStyles(doc, sink),
      theme: EMPTY_THEME,
      settings: DEFAULT_SETTINGS,
    };
    expect(resolveParaProps(ctx, pPr('<w:pStyle w:val="A"/>')).justification).toBe('right');
    expect(sink.list().map((d) => d.code)).toContain('style-missing');
  });
});

describe('制表位', () => {
  const ctx = ctxFrom(`
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
      <w:name w:val="Normal"/>
      <w:pPr><w:tabs>
        <w:tab w:val="left" w:pos="720"/>
        <w:tab w:val="right" w:pos="8306" w:leader="dot"/>
      </w:tabs></w:pPr>
    </w:style>
  `);

  it('样式与直接格式的制表位逐个合并，不是整块替换', () => {
    const p = resolveParaProps(ctx, pPr('<w:tabs><w:tab w:val="center" w:pos="4153"/></w:tabs>'));
    expect(p.tabs.map((t) => t.pos)).toEqual([720, 4153, 8306]); // 按 pos 升序
    expect(p.tabs[1]?.alignment).toBe('center');
  });

  it('同一个 pos 上后来者覆盖', () => {
    const p = resolveParaProps(ctx, pPr('<w:tabs><w:tab w:val="decimal" w:pos="720"/></w:tabs>'));
    expect(p.tabs.filter((t) => t.pos === 720)).toHaveLength(1);
    expect(p.tabs[0]?.alignment).toBe('decimal');
  });

  it('w:val="clear" 删掉继承来的那个，自己不留下任何东西', () => {
    const p = resolveParaProps(ctx, pPr('<w:tabs><w:tab w:val="clear" w:pos="720"/></w:tabs>'));
    expect(p.tabs.map((t) => t.pos)).toEqual([8306]);
    // clear 只在级联时有意义，结果里不该再出现
    expect(p.tabs.some((t) => t.alignment === 'clear')).toBe(false);
  });

  it('前导符原样保留 —— 目录那排点靠它', () => {
    expect(resolveParaProps(ctx, pPr('')).tabs[1]?.leader).toBe('dot');
  });

  it('没有制表位时是空数组，布局层据此只走 defaultTabStop', () => {
    expect(resolveParaProps(ctxFrom(''), pPr('')).tabs).toEqual([]);
  });
});
