/**
 * 级联备忘（`ResolveCache`）：没改的段落跨事务复用同一个冻结结果。
 *
 * 判据只有一条：**带备忘与不带备忘的级联结果处处相等**。复用是纯粹的性能手段，
 * 任何一处「该重算却复用了」都是正确性 bug —— 所以每个用例都拿无备忘的结果兜底，
 * 再额外看一眼「该复用的确实复用了」（不然备忘等于没有，性能那一半就白做了）。
 */
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { CascadeContext } from './cascade.ts';
import type { FieldHyperlink } from './fields.ts';
import type { Body, NodeId, ResolvedBody, ResolvedParagraph } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { parseNumbering } from './numbering.ts';
import { parseBody } from './parse-body.ts';
import type { DocPosition } from './position.ts';
import { createResolveCache, resolveBody } from './resolve-body.ts';
import { DEFAULT_SETTINGS } from './settings.ts';
import { parseStyles } from './styles.ts';
import { createTextEditor } from './text-transaction.ts';
import { EMPTY_THEME } from './theme.ts';

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const STYLES = `
  <w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/>
    <w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>
  </w:style>
  <w:style w:type="table" w:styleId="grid"><w:name w:val="Table Grid"/>
    <w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:tblStylePr>
  </w:style>`;

const NUMBERING = `
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`;

const item = (text: string) =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const plain = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string) => `<w:tc>${plain(text)}</w:tc>`;

const DOC = `
  ${plain('标题')}
  ${item('甲')}${item('乙')}
  ${plain('正文')}
  <w:tbl>
    <w:tblPr><w:tblStyle w:val="grid"/><w:tblLook w:firstRow="1"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
    <w:tr>${cell('表头')}${cell('表头二')}</w:tr>
    <w:tr>${cell('一')}${cell('二')}</w:tr>
  </w:tbl>
  ${plain('尾')}`;

function setup() {
  const sink = createDiagnosticSink();
  const numbering = parseNumbering(parseXml(`<w:numbering ${W_NS}>${NUMBERING}</w:numbering>`), sink);
  const ctx: CascadeContext = {
    styles: parseStyles(parseXml(`<w:styles ${W_NS}>${STYLES}</w:styles>`), sink),
    theme: EMPTY_THEME,
    settings: DEFAULT_SETTINGS,
    numbering,
  };
  const body: Body = {
    ...parseBody(parseXml(`<w:document ${W_NS}><w:body>${DOC}<w:sectPr/></w:body></w:document>`), sink),
    numbering,
  };
  const editor = createTextEditor(body);
  const cache = createResolveCache();
  /** 带备忘的一趟，顺手断言与无备忘的那一趟处处相等 */
  const resolve = (hyperlinks?: ReadonlyMap<NodeId, FieldHyperlink>): ResolvedBody => {
    const base = hyperlinks === undefined ? {} : { hyperlinks };
    const cached = resolveBody(ctx, editor.body, { ...base, cache });
    expect(cached).toEqual(resolveBody(ctx, editor.body, base));
    return cached;
  };
  const at = (text: string): DocPosition => {
    for (const p of walkParagraphs(editor.body))
      for (const r of p.runs)
        if (r.content.some((c) => c.kind === 'text' && c.text === text))
          return { nodeId: r.id, contentIndex: 0, offset: 0 };
    throw new Error(`找不到「${text}」`);
  };
  return { ctx, editor, cache, resolve, at };
}

const paragraphs = (body: ResolvedBody): ResolvedParagraph[] => [...walkParagraphs(body)];
const byText = (body: ResolvedBody, text: string): ResolvedParagraph => {
  const found = paragraphs(body).find((p) =>
    p.runs.some((r) => r.content.some((c) => c.kind === 'text' && c.text.includes(text))),
  );
  if (found === undefined) throw new Error(`找不到「${text}」`);
  return found;
};

describe('级联备忘', () => {
  it('只改一段：其余没有编号的段落复用同一个冻结对象，改的那段重算', () => {
    const { editor, resolve, at } = setup();
    const before = resolve();
    editor.tx((t) => {
      t.insertText(at('正文'), '新');
      return undefined;
    });
    const after = resolve();
    expect(byText(after, '标题')).toBe(byText(before, '标题'));
    expect(byText(after, '尾')).toBe(byText(before, '尾'));
    expect(byText(after, '新正文')).not.toBe(byText(before, '正文'));
    expect(Object.isFrozen(byText(after, '标题').runs[0]?.props)).toBe(true);
  });

  it('编号段落每趟都重算：前面插一项，后面的「第几」跟着变', () => {
    const { editor, resolve, at } = setup();
    const before = resolve();
    expect(byText(before, '乙').props.numbering.label?.text).toBe('2.');
    editor.tx((t) => {
      const p = t.splitParagraph(at('甲'));
      t.insertText(p, '丙');
      return undefined;
    });
    // 「乙」这一段的源节点没动，但它现在是第 3 项
    expect(byText(resolve(), '乙').props.numbering.label?.text).toBe('3.');
  });

  it('表格插一行到最上面：原表头的格不再命中 firstRow，格内段落重算', () => {
    const { editor, resolve, at } = setup();
    const before = resolve();
    expect(byText(before, '表头').runs[0]?.props.bold).toBe(true);
    editor.tx((t) => {
      t.insertRow(at('表头'), 'above');
      return undefined;
    });
    const after = resolve();
    expect(byText(after, '表头').runs[0]?.props.bold).toBe(false);
    // 表外的段落照旧复用
    expect(byText(after, '尾')).toBe(byText(before, '尾'));
  });

  it('HYPERLINK 域给的链接变了就重算，没变就复用', () => {
    const { resolve } = setup();
    const runId = byText(resolve(), '尾').runs[0]?.id as NodeId;
    const linked = resolve(new Map([[runId, { url: 'https://a.example' }]]));
    expect(byText(linked, '尾').runs[0]?.hyperlink).toEqual({ url: 'https://a.example' });
    const moved = resolve(new Map([[runId, { url: 'https://b.example' }]]));
    expect(byText(moved, '尾').runs[0]?.hyperlink).toEqual({ url: 'https://b.example' });
    const same = resolve(new Map([[runId, { url: 'https://b.example' }]]));
    expect(byText(same, '尾')).toBe(byText(moved, '尾'));
    expect(byText(resolve(), '尾').runs[0]?.hyperlink).toBeUndefined();
  });

  it('树上补了样式：整份作废重算（样式能改任何一段）', () => {
    const { editor, resolve, at } = setup();
    const before = resolve();
    editor.tx((t) => {
      const id = t.addStyle({ id: 'custom', name: '自定义', paraProps: {}, runProps: { bold: true } });
      t.setParagraphProps({ start: at('尾'), end: at('尾') }, { styleId: id });
      return undefined;
    });
    const after = resolve();
    expect(byText(after, '标题')).not.toBe(byText(before, '标题'));
    expect(byText(after, '尾').runs[0]?.props.bold).toBe(true);
  });

  it('撤销回到旧快照：旧段落对象重新出现时直接命中', () => {
    const { editor, resolve, at } = setup();
    const before = resolve();
    editor.tx((t) => {
      t.insertText(at('正文'), '新');
      return undefined;
    });
    resolve();
    editor.undo();
    expect(byText(resolve(), '正文')).toBe(byText(before, '正文'));
  });

  it('可写的源段落不进备忘 —— 原地改过的也能看见', () => {
    const { ctx, cache } = setup();
    const body = parseBody(
      parseXml(`<w:document ${W_NS}><w:body>${plain('可写')}<w:sectPr/></w:body></w:document>`),
      createDiagnosticSink(),
    );
    const first = resolveBody(ctx, body, { cache });
    const run = [...walkParagraphs(body)][0]?.runs[0];
    if (run === undefined) throw new Error('缺 run');
    run.props = { ...run.props, bold: true };
    const second = resolveBody(ctx, body, { cache });
    expect(paragraphs(second)[0]).not.toBe(paragraphs(first)[0]);
    expect(paragraphs(second)[0]?.runs[0]?.props.bold).toBe(true);
  });
});
