import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { CascadeContext } from './cascade.ts';
import type { ResolvedParagraph } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { EMPTY_NUMBERING } from './numbering.ts';
import { rangeOfNode } from './order.ts';
import { parseBody } from './parse-body.ts';
import { resolveBody } from './resolve-body.ts';
import { DEFAULT_SETTINGS } from './settings.ts';
import { extendStyleSheet, parseStyles } from './styles.ts';
import { builtinStyleDefinition } from './styles-edit.ts';
import { createTextEditor } from './text-transaction.ts';
import { EMPTY_THEME } from './theme.ts';

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const STYLES = `
  <w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>
    <w:rPr><w:sz w:val="21"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="My Heading"/></w:style>`;

function ctx(): CascadeContext {
  return {
    styles: parseStyles(parseXml(`<w:styles ${W_NS}>${STYLES}</w:styles>`), createDiagnosticSink()),
    theme: EMPTY_THEME,
    settings: DEFAULT_SETTINGS,
    numbering: EMPTY_NUMBERING,
  };
}

function editable() {
  const context = ctx();
  const body = parseBody(
    parseXml(`<w:document><w:body><w:p><w:r><w:t>标题</w:t></w:r></w:p><w:p/></w:body></w:document>`),
    createDiagnosticSink(),
  );
  const editor = createTextEditor(body);
  const paragraphs = () => [...walkParagraphs(resolveBody(context, editor.body))] as ResolvedParagraph[];
  const start = (i: number) => {
    const p = [...walkParagraphs(editor.body)][i];
    const range = p && rangeOfNode(p);
    if (!range) throw new Error('没有段落');
    return range.start;
  };
  return { context, editor, paragraphs, start };
}

describe('内建样式定义', () => {
  it('id 撞上文档里的自定义样式时加序号，basedOn / next 指向正文', () => {
    const sheet = ctx().styles;
    const def = builtinStyleDefinition('heading 1', (id) => sheet.byId(id) !== undefined, 'a');
    expect(def).toMatchObject({ id: 'Heading1_1', name: 'heading 1', basedOn: 'a', next: 'a' });
    expect(def.runProps).toMatchObject({ bold: true, size: 440 });
    expect(def.paraProps).toMatchObject({ keepNext: true, outlineLevel: 0 });
    // 没有默认段落样式时不写 basedOn / next（写空串会成悬空引用）
    expect(builtinStyleDefinition('heading 2', () => false, '')).not.toHaveProperty('basedOn');
  });

  it('扩展表：新样式沿着 basedOn 继承正文，原表不变，同一份定义复用同一张表', () => {
    const base = ctx().styles;
    const added = [builtinStyleDefinition('heading 3', () => false, 'a')];
    const sheet = extendStyleSheet(base, added);
    expect(sheet.chainOf('Heading3').map((s) => s.id)).toEqual(['a', 'Heading3']);
    expect(sheet.byId('Heading3')?.next).toBe('a');
    expect(base.byId('Heading3')).toBeUndefined();
    expect(extendStyleSheet(base, added)).toBe(sheet);
    expect(extendStyleSheet(base, [])).toBe(base);
  });
});

describe('事务 addStyle', () => {
  it('新增定义并引用，级联用上它；撤销连定义一起回退，重做恢复', () => {
    const { editor, paragraphs, start } = editable();
    const def = builtinStyleDefinition('heading 2', () => false, 'a');
    editor.tx((t) => {
      const id = t.addStyle(def);
      t.setParagraphProps({ start: start(0), end: start(0) }, { styleId: id });
    });
    expect(editor.body.styles?.map((s) => s.id)).toEqual(['Heading2']);
    expect(paragraphs()[0]?.props).toMatchObject({ styleId: 'Heading2', keepNext: true, outlineLevel: 1 });
    expect(paragraphs()[0]?.runs[0]?.props).toMatchObject({ bold: true, size: 320 });
    // 另一段没有引用它，照旧是正文
    expect(paragraphs()[1]?.props.styleId).toBe('a');
    editor.undo();
    expect(editor.body.styles).toBeUndefined();
    expect(paragraphs()[0]?.props.styleId).toBe('a');
    editor.redo();
    expect(editor.body.styles?.[0]).toEqual(def);
  });

  it('只加定义不引用是无修改；同一 id 加两次整次回滚', () => {
    const { editor, start } = editable();
    const def = builtinStyleDefinition('heading 1', () => false, 'a');
    expect(editor.tx((t) => void t.addStyle(def))).toBeUndefined();
    expect(editor.body.styles).toBeUndefined();
    expect(() =>
      editor.tx((t) => {
        t.addStyle(def);
        t.addStyle(def);
        t.setParagraphProps({ start: start(0), end: start(0) }, { styleId: def.id });
      }),
    ).toThrow();
    expect(editor.body.styles).toBeUndefined();
    expect(editor.canUndo).toBe(false);
  });
});
