/**
 * 拿真实公文 `gongwen-01.docx` 走完整条链：解包 → 样式表 + 主题 → 级联。
 *
 * 期望值不是从代码推的，是从文件里实际写着的东西反推的（见每个断言旁的注释）。
 * 这是 Phase 1 DoD「解析后属性树与 Word 显示格式面板一致」的第一批抽查点。
 */
import { readFileSync } from 'node:fs';
import { createDiagnosticSink, twipsToPt } from '@uw/core';
import type { XmlElement } from '@uw/ooxml';
import { child, children, OpcPackage } from '@uw/ooxml';
import { beforeAll, describe, expect, it } from 'vitest';
import type { CascadeContext } from './cascade.ts';
import { resolveParaProps, resolveRunProps } from './cascade.ts';
import type { LoadedDocument } from './load.ts';
import { loadCascadeContext, loadDocument } from './load.ts';
import { paragraphText, walkParagraphs } from './nodes.ts';
import { parseParaProps, parseRunProps } from './parse-props.ts';
import { DEFAULT_SETTINGS } from './settings.ts';
import { parseStyles } from './styles.ts';
import { resolveThemeFont } from './theme.ts';

const FIXTURE = new URL('../../../apps/fidelity/fixtures/gongwen-01.docx', import.meta.url);

let pkg: OpcPackage;
let ctx: CascadeContext;
let doc: LoadedDocument;
let paragraphs: XmlElement[];
const sink = createDiagnosticSink();

beforeAll(() => {
  pkg = OpcPackage.open(new Uint8Array(readFileSync(FIXTURE)));
  ctx = loadCascadeContext(pkg, sink);
  doc = loadDocument(pkg, sink);
  const body = child(pkg.xml(pkg.mainDocumentPartName()).root, 'w:body');
  paragraphs = body === undefined ? [] : children(body, 'w:p');
});

/** 第 n 段的（段落属性，第一个 run 的字符属性） */
function para(n: number): { p: XmlElement | undefined; r: XmlElement | undefined } {
  const el = paragraphs[n];
  if (el === undefined) return { p: undefined, r: undefined };
  const firstRun = children(el, 'w:r')[0];
  return { p: child(el, 'w:pPr'), r: firstRun && child(firstRun, 'w:rPr') };
}

describe('样式表', () => {
  it('33 个样式，默认段落样式是 Normal（styleId 为 "a"）', () => {
    expect(ctx.styles.all()).toHaveLength(33);
    const id = ctx.styles.defaultParagraphStyleId();
    expect(id).toBe('a');
    expect(ctx.styles.byId(id)?.name).toBe('Normal');
  });

  it('heading 1 的 basedOn 链是 [Normal, heading 1]', () => {
    expect(ctx.styles.chainOf('1').map((s) => s.name)).toEqual(['Normal', 'heading 1']);
  });

  it('这份文档的样式表没有环，也没有悬空引用', () => {
    for (const s of ctx.styles.all()) ctx.styles.chainOf(s.id);
    expect(sink.list().filter((d) => d.code.startsWith('style-'))).toEqual([]);
  });
});

describe('主题字体', () => {
  it('minorHAnsi → 等线（a:latin）', () => {
    expect(resolveThemeFont(ctx.theme, 'minorHAnsi')).toBe('等线');
    expect(resolveThemeFont(ctx.theme, 'majorHAnsi')).toBe('等线 Light');
  });

  it('minorEastAsia 的 a:ea 是空串，必须回退到 script="Hans"', () => {
    // 文件里就是 <a:ea typeface=""/> —— 直接取会得到空字体名，整份公文的中文字体没了
    expect(ctx.theme.minor.eastAsia).toBe('');
    expect(resolveThemeFont(ctx.theme, 'minorEastAsia')).toBe('等线');
    expect(resolveThemeFont(ctx.theme, 'majorEastAsia')).toBe('等线 Light');
  });

  it('按语言选脚本回退', () => {
    expect(resolveThemeFont(ctx.theme, 'minorEastAsia', 'ja-JP')).toBe('游明朝');
    expect(resolveThemeFont(ctx.theme, 'minorEastAsia', 'zh-TW')).toBe('新細明體');
  });
});

describe('docDefaults 级联', () => {
  it('默认字号 10.5pt（五号），默认中文字体经主题解析成等线', () => {
    const r = resolveRunProps(ctx, undefined, undefined);
    expect(twipsToPt(r.size)).toBe(10.5); // <w:sz w:val="21"/>
    expect(r.fonts.eastAsia).toBe('等线'); // eastAsiaTheme="minorEastAsia" → a:ea 空 → Hans
    expect(r.fonts.ascii).toBe('等线'); // asciiTheme="minorHAnsi" → a:latin
    expect(r.langEastAsia).toBe('zh-CN');
  });

  it('Normal 样式关掉了孤行控制、并且两端对齐', () => {
    // <w:style w:styleId="a"><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/>
    const p = resolveParaProps(ctx, undefined);
    expect(p.widowControl).toBe(false);
    expect(p.justification).toBe('both');
  });
});

describe('正文段落抽查', () => {
  it('标题段：居中、段后 24pt、黑体二号（22pt）', () => {
    const { p, r } = para(0);
    const rp = resolveParaProps(ctx, parseParaProps(p));
    const rr = resolveRunProps(ctx, parseParaProps(p), parseRunProps(r));
    expect(rp.justification).toBe('center');
    expect(twipsToPt(rp.spacing.after)).toBe(24); // w:after="480"
    expect(rr.fonts.eastAsia).toBe('黑体');
    expect(rr.fonts.ascii).toBe('Times New Roman');
    expect(rr.fonts.hint).toBe('eastAsia');
    expect(twipsToPt(rr.size)).toBe(22); // w:sz="44"
  });

  it('主送机关段：左对齐（直接格式压过 Normal 的 both）、仿宋三号（16pt）', () => {
    const { p, r } = para(1);
    expect(resolveParaProps(ctx, parseParaProps(p)).justification).toBe('left');
    const rr = resolveRunProps(ctx, parseParaProps(p), parseRunProps(r));
    expect(rr.fonts.eastAsia).toBe('仿宋');
    expect(twipsToPt(rr.size)).toBe(16); // w:sz="32"
  });

  it('正文段：首行缩进 2 字符，且字符单位与 twips 版本同时存在', () => {
    const rp = resolveParaProps(ctx, parseParaProps(para(2).p));
    // <w:ind w:firstLineChars="200" w:firstLine="640"/>
    expect(rp.indent.firstLineChars).toBe(200); // 1/100 字符 → 2 字符
    expect(rp.indent.firstLine).toBe(640);
    // 640 twips = 32pt = 2 × 16pt，两者此刻恰好一致；字号一变就只有前者还对，
    // 所以布局层必须优先用 firstLineChars
    expect(twipsToPt(rp.indent.firstLine)).toBe(2 * 16);
  });

  it('整份文档的段落都能解析出 Resolved 属性，且可结构化克隆', () => {
    expect(paragraphs.length).toBe(7);
    for (const el of paragraphs) {
      const rp = resolveParaProps(ctx, parseParaProps(child(el, 'w:pPr')));
      expect(structuredClone(rp)).toEqual(rp);
      expect(rp.spacing.lineRule).toBeTypeOf('string');
    }
  });

  it('整个流程一条诊断都不产生 —— 这份文档我们全认识', () => {
    expect(sink.list()).toEqual([]);
  });
});

describe('正文节点树（loadDocument 全链路）', () => {
  it('一节，A4 纵向，页边距与行网格与文件里写的一致', () => {
    const s = doc.body.sections[0];
    expect(doc.body.sections).toHaveLength(1);
    expect(s?.props.page).toEqual({ width: 11906, height: 16838, orientation: 'portrait' });
    // <w:pgMar w:top="2098" w:right="1474" w:bottom="1984" w:left="1587" w:header="851" w:footer="992" w:gutter="0"/>
    expect(s?.props.margin).toEqual({
      top: 2098,
      right: 1474,
      bottom: 1984,
      left: 1587,
      header: 851,
      footer: 992,
      gutter: 0,
    });
    // <w:docGrid w:type="lines" w:linePitch="579"/> —— 579 twips = 28.95pt，公文的每页行数由它定
    expect(s?.props.docGrid).toEqual({ type: 'lines', linePitch: 579, charSpace: 0 });
    // <w:cols w:space="425"/> 没写 w:num，就是单栏
    expect(s?.props.columns).toBe(1);
  });

  it('7 个段落，文字与 XML 里的一字不差（跨 run 的中英混排也要接得上）', () => {
    const texts = [...walkParagraphs(doc.body)].map(paragraphText);
    expect(texts).toHaveLength(7);
    expect(texts[0]).toBe('关于进一步加强文档排版引擎保真度验证工作的通知');
    expect(texts[1]).toBe('各有关单位：');
    // 这一段被 Word 切成 15 个 run（每遇中英切换就断一次），拼回来必须严丝合缝
    expect(texts[2]).toContain('自 2026 年起，所有版式输出均须以 Word 导出的 PDF 坐标为准');
    expect(texts[6]).toBe('2026 年 8 月 13 日');
  });

  it('resolved 树与可编辑树同形：id、顺序、层级都一一对应', () => {
    const src = [...walkParagraphs(doc.body)];
    const out = [...walkParagraphs(doc.resolved)];
    expect(out.map((p) => p.id)).toEqual(src.map((p) => p.id));
    expect(out.map((p) => p.runs.map((r) => r.id))).toEqual(src.map((p) => p.runs.map((r) => r.id)));
    expect(out.map(paragraphText)).toEqual(src.map(paragraphText));
  });

  it('resolved 树上的属性就是级联结果 —— 标题段黑体二号居中', () => {
    const p = [...walkParagraphs(doc.resolved)][0];
    expect(p?.props.justification).toBe('center');
    expect(p?.runs[0]?.props.fonts.eastAsia).toBe('黑体');
    expect(twipsToPt(p?.runs[0]?.props.size ?? 0)).toBe(22);
    // 正文段落没写 snapToGrid，Word 默认是开的 —— 公文的行网格靠它生效
    expect(p?.props.snapToGrid).toBe(true);
  });

  it('resolved 树整棵可结构化克隆 —— Worker 边界的门票（原则 1.1）', () => {
    expect(structuredClone(doc.resolved)).toEqual(doc.resolved);
  });

  it('字体表与设置跟着 loadDocument 一起出来', () => {
    // 「黑体」→ SimHei：非中文系统上查字体全靠这一步
    expect(doc.fonts.byName.黑体?.altName).toBe('SimHei');
    expect(doc.cascade.settings.defaultTabStop).toBe(420);
    // 这份文档没有 numbering.xml
    expect(doc.numbering).toEqual({ abstract: {}, instances: {} });
  });

  it('themeFontLang 驱动主题东亚字体的回退；run 自己的 w:lang 优先级更高', () => {
    // 这份文档主题里 <a:ea typeface=""/> 是空的，东亚字体靠 script 回退，选哪个 script
    // 由 themeFontLang.eastAsia 决定。用空样式表隔离出这一条线：
    const langCtx = (eastAsia: string): CascadeContext => ({
      styles: parseStyles(undefined, createDiagnosticSink()),
      theme: doc.cascade.theme,
      settings: { ...DEFAULT_SETTINGS, themeFontLang: { latin: '', eastAsia, bidi: '' } },
      numbering: doc.cascade.numbering,
    });
    const themeRef = { fontThemes: { eastAsia: 'minorEastAsia' } };
    expect(resolveRunProps(langCtx('ja-JP'), undefined, themeRef).fonts.eastAsia).toBe('游明朝');
    expect(resolveRunProps(langCtx('zh-TW'), undefined, themeRef).fonts.eastAsia).toBe('新細明體');
    // 两边都没写时兜底 zh-CN —— 这个库的定位是中文公文
    expect(resolveRunProps(langCtx(''), undefined, themeRef).fonts.eastAsia).toBe('等线');

    // 而这份文档的 docDefaults 里写了 <w:lang w:eastAsia="zh-CN"/>，run 级的语言压过文档级设置：
    // 把 themeFontLang 改成日文也不该动摇它
    const ja: CascadeContext = {
      ...doc.cascade,
      settings: {
        ...doc.cascade.settings,
        themeFontLang: { latin: 'en-US', eastAsia: 'ja-JP', bidi: '' },
      },
    };
    expect(resolveRunProps(ja, undefined, undefined).fonts.eastAsia).toBe('等线');
  });

  it('解析整份文档仍然一条诊断都不产生', () => {
    expect(sink.list()).toEqual([]);
  });
});
