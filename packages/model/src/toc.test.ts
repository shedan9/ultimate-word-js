/**
 * 目录重新生成：`planTableOfContents` 挑标题拼草稿，事务 `addBookmark` / `replaceFieldResult` 落到树上。
 * 回写（新书签、新条目段落怎么写回 XML）在 serialize 的 docx.test.ts，接线与页码在门面的 facade.test.ts。
 * 条目的样子照 Word 生成的目录写，**没有真值样本**（见 toc.ts 文件头）。
 */
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { CascadeContext } from './cascade.ts';
import { scanFields } from './fields.ts';
import type { Body, Paragraph, RunContent } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { EMPTY_NUMBERING } from './numbering.ts';
import { parseBody } from './parse-body.ts';
import { resolveBody } from './resolve-body.ts';
import { DEFAULT_SETTINGS } from './settings.ts';
import { parseStyles } from './styles.ts';
import { paragraphStyleNames } from './styles-edit.ts';
import { createTextEditor } from './text-transaction.ts';
import { EMPTY_THEME } from './theme.ts';
import type { TocPlan, TocPlanOptions } from './toc.ts';
import { planTableOfContents, TOC_EMPTY_TEXT, tocFields } from './toc.ts';

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** 中文版 Word 的标题 id 是数字，名字是英文 —— 按大纲级别认，不按 id */
const CTX: CascadeContext = {
  styles: parseStyles(
    parseXml(
      `<w:styles ${W_NS}>` +
        '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>' +
        '<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="4"><w:name w:val="heading 4"/><w:pPr><w:outlineLvl w:val="3"/></w:pPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="11"><w:name w:val="toc 1"/></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Memo"><w:name w:val="备忘"/></w:style>' +
        '</w:styles>',
    ),
    createDiagnosticSink(),
  ),
  theme: EMPTY_THEME,
  settings: DEFAULT_SETTINGS,
  numbering: EMPTY_NUMBERING,
};

const OPTS: TocPlanOptions = {
  styleNames: paragraphStyleNames(CTX.styles),
  normalStyleId: 'a',
  pageText: () => '7',
};

const SECT =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const fld = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const instr = (text: string) => `<w:r><w:instrText xml:space="preserve"> ${text} </w:instrText></w:r>`;
const heading = (style: string, text: string, bookmark = '') =>
  `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${bookmark}<w:r><w:t>${text}</w:t></w:r></w:p>`;

/** 一个「陈旧」的目录：只有一条「旧标题」，end 单独一段 */
function doc(tocSwitches: string, rest: string): Body {
  const xml =
    `<w:p><w:pPr><w:pStyle w:val="11"/></w:pPr>${fld('begin')}${instr(`TOC ${tocSwitches}`)}${fld('separate')}` +
    `<w:r><w:t>旧标题</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:p>` +
    `<w:p>${fld('end')}</w:p>${rest}${SECT}`;
  return parseBody(
    parseXml(`<w:document ${W_NS}><w:body>${xml}</w:body></w:document>`),
    createDiagnosticSink(),
  );
}

function plan(body: Body): TocPlan {
  const fields = scanFields(body);
  const toc = tocFields(fields)[0];
  if (toc === undefined) throw new Error('没有目录');
  const result = planTableOfContents(resolveBody(CTX, body), body, fields, toc, OPTS);
  if (!result.ok) throw new Error(result.reason);
  return result.plan;
}

/** 草稿里一段的文字：制表位写成 →，PAGEREF 指令写成 {…} */
function draftText(content: readonly RunContent[]): string {
  return content
    .map((c) =>
      c.kind === 'text'
        ? c.text
        : c.kind === 'tab'
          ? '→'
          : c.kind === 'fieldInstruction'
            ? `{${c.text.trim()}}`
            : '',
    )
    .join('');
}

function text(p: Paragraph): string {
  return draftText(p.runs.flatMap((r) => r.content));
}

describe('planTableOfContents', () => {
  const REST =
    heading('1', '第一章 总则', '<w:bookmarkStart w:id="0" w:name="_Toc100000123"/>') +
    heading('2', '一、适用范围') +
    heading('4', '四级标题不进') +
    heading('a', '正文不进') +
    heading('1', '   ') +
    heading('1', '第二章 附则');

  it('按大纲级别挑 \\o 范围内的标题，沿用已有 _Toc 书签、给没有的补新的，跳过空标题', () => {
    const p = plan(doc('\\o "1-3" \\h \\z \\u', REST));
    expect(p.entries).toBe(3);
    expect(p.paragraphs.map((d) => draftText(d.runs.flatMap((r) => r.content)))).toEqual([
      '第一章 总则→{PAGEREF _Toc100000123 \\h}7',
      '一、适用范围→{PAGEREF _Toc100000124 \\h}7',
      '第二章 附则→{PAGEREF _Toc100000125 \\h}7',
    ]);
    expect(p.bookmarks.map((b) => b.name)).toEqual(['_Toc100000124', '_Toc100000125']);
    // \h：整条都是跳到标题书签的超链接
    expect(p.paragraphs[1]?.runs.every((r) => r.hyperlink?.anchor === '_Toc100000124')).toBe(true);
  });

  it('条目用 toc N 样式（按名字认，缺的补一份定义），右对齐点前导制表位停在版心宽', () => {
    const p = plan(doc('\\o "1-3"', REST));
    expect(p.paragraphs[0]?.props).toEqual({
      styleId: '11',
      tabs: [{ pos: 11906 - 3600, alignment: 'right', leader: 'dot' }],
    });
    expect(p.styles).toHaveLength(1);
    expect(p.styles[0]).toMatchObject({
      id: 'TOC2',
      name: 'toc 2',
      basedOn: 'a',
      paraProps: { indent: { left: 420 } },
    });
    expect(p.paragraphs[1]?.props.styleId).toBe('TOC2');
    // 没有 \h 就没有超链接
    expect(p.paragraphs[0]?.runs.some((r) => r.hyperlink !== undefined)).toBe(false);
  });

  it('\\n 去掉页码与制表位，\\t 按样式名点名级别，\\p 换掉分隔符', () => {
    const noPages = plan(doc('\\o "1-2" \\n "2-2"', REST));
    expect(noPages.paragraphs.map((d) => draftText(d.runs.flatMap((r) => r.content)))[1]).toBe(
      '一、适用范围',
    );
    expect(noPages.paragraphs[1]?.props.tabs).toBeUndefined();

    const named = plan(doc('\\t "备忘,2"', `${heading('Memo', '备忘录')}${heading('1', '标题不点名')}`));
    expect(named.paragraphs.map((d) => draftText(d.runs.flatMap((r) => r.content)))).toEqual([
      '备忘录→{PAGEREF _Toc100000000 \\h}7',
    ]);

    const dotted = plan(doc('\\o "1-1" \\p "……"', heading('1', '甲')));
    expect(draftText(dotted.paragraphs[0]?.runs.flatMap((r) => r.content) ?? [])).toBe(
      '甲……{PAGEREF _Toc100000000 \\h}7',
    );
  });

  it('没有标题时写「未找到目录项。」；图表目录（\\c）不更新', () => {
    const empty = plan(doc('\\o "1-3"', heading('a', '只有正文')));
    expect(empty.entries).toBe(0);
    expect(draftText(empty.paragraphs[0]?.runs.flatMap((r) => r.content) ?? [])).toBe(TOC_EMPTY_TEXT);

    const body = doc('\\h \\z \\c "图"', REST);
    const fields = scanFields(body);
    const toc = tocFields(fields)[0];
    if (toc === undefined) throw new Error('没有目录');
    expect(planTableOfContents(resolveBody(CTX, body), body, fields, toc, OPTS)).toMatchObject({ ok: false });
  });

  it('目录自己的条目不再收进目录（toc 样式的段落带大纲级别也一样）', () => {
    const body = doc('\\u', heading('1', '唯一标题'));
    // 把旧条目那一段也设成标题 1：它在结果区里，不能被收
    const p = plan({
      ...body,
      sections: body.sections.map((s) => ({
        ...s,
        blocks: s.blocks.map((b, i) =>
          i === 0 && b.kind === 'paragraph' ? { ...b, props: { ...b.props, styleId: '1' } } : b,
        ),
      })),
    });
    expect(p.entries).toBe(1);
  });
});

describe('事务：addBookmark + replaceFieldResult', () => {
  const REST = heading('1', '第一章') + heading('2', '第一节') + heading('1', '第二章');

  function update(body: Body) {
    const editor = createTextEditor(body);
    const p = plan(editor.body);
    const change = editor.tx((t) => {
      for (const b of p.bookmarks) t.addBookmark(b.paragraphId, b.name);
      t.replaceFieldResult(p.field, p.paragraphs);
    });
    return { editor, change };
  }

  it('换掉结果区：界桩与指令原样、首段沿用 separate 那一段、end 那一段留在最后，撤销整个退回', () => {
    const body = doc('\\o "1-3" \\h', REST);
    const before = [...walkParagraphs(body)];
    const { editor, change } = update(body);
    const after = [...walkParagraphs(editor.body)];
    expect(change).toBeDefined();
    // 三条目录 + end 那一段 + 三个标题
    expect(after).toHaveLength(7);
    expect(after.slice(0, 4).map(text)).toEqual([
      '{TOC \\o "1-3" \\h}第一章→{PAGEREF _Toc100000000 \\h}7',
      '第一节→{PAGEREF _Toc100000001 \\h}7',
      '第二章→{PAGEREF _Toc100000002 \\h}7',
      '',
    ]);
    // 首段与 end 段的 id 不变（回写吐回原元素），界桩 run 也是原来那几个
    expect(after[0]?.id).toBe(before[0]?.id);
    expect(after[3]).toBe(editor.body.sections[0]?.blocks[3]);
    expect(after[3]?.id).toBe(before[1]?.id);
    expect(after[0]?.runs.slice(0, 3).map((r) => r.id)).toEqual(before[0]?.runs.slice(0, 3).map((r) => r.id));
    // 书签补在标题上；新结果仍是一个配得上对的完整域，嵌套的 PAGEREF 各自成域
    expect(after[4]?.bookmarks).toEqual(['_Toc100000000']);
    const fields = scanFields(editor.body);
    expect(fields.filter((f) => f.instruction.type === 'PAGEREF')).toHaveLength(3);
    expect(tocFields(fields)[0]?.end?.paragraphId).toBe(before[1]?.id);

    editor.undo();
    expect([...walkParagraphs(editor.body)].map(text)).toEqual(before.map(text));
    expect([...walkParagraphs(editor.body)][4]?.bookmarks).toBeUndefined();
  });

  it('end 在最后一条的末尾时，最后一条沿用那一段；结果区在同一段里也能换', () => {
    const tail =
      `<w:p><w:pPr><w:pStyle w:val="11"/></w:pPr>${fld('begin')}${instr('TOC \\o "1-3"')}${fld('separate')}` +
      `<w:r><w:t>甲</w:t></w:r></w:p><w:p><w:r><w:t>乙</w:t></w:r>${fld('end')}</w:p>`;
    const xml = `${tail}${REST}${SECT}`;
    const body = parseBody(
      parseXml(`<w:document ${W_NS}><w:body>${xml}</w:body></w:document>`),
      createDiagnosticSink(),
    );
    const before = [...walkParagraphs(body)];
    const { editor } = update(body);
    const after = [...walkParagraphs(editor.body)];
    expect(after.slice(0, 3).map(text)).toEqual([
      '{TOC \\o "1-3"}第一章→{PAGEREF _Toc100000000 \\h}7',
      '第一节→{PAGEREF _Toc100000001 \\h}7',
      '第二章→{PAGEREF _Toc100000002 \\h}7',
    ]);
    expect(after[2]?.id).toBe(before[1]?.id);
    expect(after).toHaveLength(6);

    const single =
      `<w:p>${fld('begin')}${instr('TOC \\o "1-3"')}${fld('separate')}<w:r><w:t>旧</w:t></w:r>${fld('end')}` +
      `<w:r><w:t>尾巴</w:t></w:r></w:p>${REST}${SECT}`;
    const one = update(
      parseBody(
        parseXml(`<w:document ${W_NS}><w:body>${single}</w:body></w:document>`),
        createDiagnosticSink(),
      ),
    );
    const got = [...walkParagraphs(one.editor.body)].map(text);
    expect(got.slice(0, 3)).toEqual([
      '{TOC \\o "1-3"}第一章→{PAGEREF _Toc100000000 \\h}7',
      '第一节→{PAGEREF _Toc100000001 \\h}7',
      '第二章→{PAGEREF _Toc100000002 \\h}7尾巴',
    ]);
  });

  it('草稿里界桩没配平、书签重名时整次回滚', () => {
    const editor = createTextEditor(doc('\\o "1-3"', REST));
    const p = plan(editor.body);
    expect(() =>
      editor.tx((t) => {
        t.replaceFieldResult(p.field, [
          { props: {}, runs: [{ content: [{ kind: 'fieldChar', charType: 'begin' }] }] },
        ]);
      }),
    ).toThrow(/配平/);
    const [first, second] = p.bookmarks;
    if (first === undefined || second === undefined) throw new Error('应当要补书签');
    expect(() =>
      editor.tx((t) => {
        t.addBookmark(first.paragraphId, 'X1');
        t.addBookmark(second.paragraphId, 'X1');
      }),
    ).toThrow(/已存在/);
    expect(editor.canUndo).toBe(false);
  });
});
