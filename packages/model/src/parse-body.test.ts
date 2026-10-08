/**
 * 正文节点树的解析。
 *
 * 用手写的最小 XML 片段，每个用例只打一个点 —— 真实文档的验证在 fixture.test.ts。
 * 这里的重点是那些**丢内容才会暴露**的分支：透明容器、修订、内容控件、合并单元格。
 */

import type { DiagnosticSink } from '@uw/core';
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { Block, Body, Paragraph, RunContent } from './nodes.ts';
import { paragraphText, walkParagraphs } from './nodes.ts';
import { parseBody, parseNotes } from './parse-body.ts';

/** 把正文片段包成一份 document.xml 再解析。命名空间不必声明 —— 这一层只按前缀名匹配 */
function parse(bodyXml: string): { body: Body; sink: DiagnosticSink } {
  const sink = createDiagnosticSink();
  const doc = parseXml(`<w:document><w:body>${bodyXml}</w:body></w:document>`, 'document.xml');
  return { body: parseBody(doc, sink), sink };
}

/** 只有一节、只关心段落时的快捷方式 */
function paras(bodyXml: string): { list: Paragraph[]; sink: DiagnosticSink } {
  const { body, sink } = parse(bodyXml);
  return { list: [...walkParagraphs(body)] as Paragraph[], sink };
}

const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';

function blockText(b: Block | undefined): string {
  return b !== undefined && b.kind === 'paragraph' ? paragraphText(b) : '';
}

describe('书签起点（PAGEREF 的目标）', () => {
  it('段内的起点记在本段上，按出现顺序；隔着超链接也算', () => {
    const { list, sink } = paras(
      `<w:p><w:bookmarkStart w:id="0" w:name="_Toc1"/><w:r><w:t>一</w:t></w:r><w:bookmarkEnd w:id="0"/>` +
        `<w:hyperlink w:anchor="x"><w:bookmarkStart w:id="1" w:name="甲"/><w:r><w:t>二</w:t></w:r></w:hyperlink></w:p>` +
        `<w:p><w:r><w:t>三</w:t></w:r></w:p>${SECT}`,
    );
    expect(list[0]?.bookmarks).toEqual(['_Toc1', '甲']);
    expect(list[1]?.bookmarks).toBeUndefined();
    expect(sink.list()).toEqual([]);
  });

  it('段落之间（body / 单元格直下）的起点归文档序里的下一段，文末剩下的归最后一段', () => {
    const { list } = paras(
      `<w:p/><w:bookmarkStart w:id="0" w:name="甲"/>` +
        `<w:tbl><w:tr><w:tc><w:bookmarkStart w:id="1" w:name="乙"/><w:p/></w:tc></w:tr></w:tbl>` +
        `<w:p/><w:bookmarkStart w:id="2" w:name="丙"/>${SECT}`,
    );
    expect(list.map((p) => p.bookmarks)).toEqual([undefined, ['甲', '乙'], ['丙']]);
  });

  it('Word 自己的 _GoBack 不收', () => {
    const { list } = paras(`<w:p><w:bookmarkStart w:id="0" w:name="_GoBack"/></w:p>${SECT}`);
    expect(list[0]?.bookmarks).toBeUndefined();
  });
});

describe('段落与 run', () => {
  it('run 里的片段按出现顺序保留，制表位不会被并进文字', () => {
    const { list } = paras(`<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>${SECT}`);
    expect(list[0]?.runs[0]?.content).toEqual<RunContent[]>([
      { kind: 'text', text: 'a' },
      { kind: 'tab' },
      { kind: 'text', text: 'b' },
    ]);
  });

  it('没有 xml:space="preserve" 时首尾空白按 XML 规矩去掉，有则原样保留', () => {
    const { list } = paras(
      `<w:p><w:r><w:t>  x  </w:t><w:t xml:space="preserve">  y  </w:t></w:r></w:p>${SECT}`,
    );
    expect(list[0]?.runs[0]?.content).toEqual<RunContent[]>([
      { kind: 'text', text: 'x' },
      { kind: 'text', text: '  y  ' },
    ]);
  });

  it('w:br 缺省是换行，w:type 才区分分页与分栏；w:cr 等同换行', () => {
    const { list } = paras(
      `<w:p><w:r><w:br/><w:br w:type="page"/><w:br w:type="column"/><w:cr/></w:r></w:p>${SECT}`,
    );
    expect(list[0]?.runs[0]?.content).toEqual<RunContent[]>([
      { kind: 'break', breakType: 'line' },
      { kind: 'break', breakType: 'page' },
      { kind: 'break', breakType: 'column' },
      { kind: 'break', breakType: 'line' },
    ]);
  });

  it('两种连字符是不同的东西，不能都当成 "-"', () => {
    const { list } = paras(`<w:p><w:r><w:noBreakHyphen/><w:softHyphen/></w:r></w:p>${SECT}`);
    expect(list[0]?.runs[0]?.content.map((c) => c.kind)).toEqual(['noBreakHyphen', 'softHyphen']);
  });

  it('w:sym 的 w:char 是十六进制码位，要解成真字符', () => {
    const { list } = paras(`<w:p><w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r></w:p>${SECT}`);
    expect(list[0]?.runs[0]?.content[0]).toEqual({
      kind: 'symbol',
      font: 'Wingdings',
      // F0FC 落在私用区（Wingdings 的对钩），源码里不塞这个字符，写码位更清楚
      char: String.fromCodePoint(0xf0fc),
    });
  });

  it('域：界桩位置与域代码都留着，结果文字照常是普通 run 文本', () => {
    const { list } = paras(
      `<w:p>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>3</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>${SECT}`,
    );
    const p = list[0] as Paragraph;
    expect(p.runs.flatMap((r) => r.content)).toEqual<RunContent[]>([
      { kind: 'fieldChar', charType: 'begin' },
      { kind: 'fieldInstruction', text: ' PAGE ' },
      { kind: 'fieldChar', charType: 'separate' },
      { kind: 'text', text: '3' },
      { kind: 'fieldChar', charType: 'end' },
    ]);
  });

  it('w:drawing 的 wp:extent 是 EMU，要换成 twips（1 英寸 = 914400 EMU = 1440 twips）', () => {
    const { list } = paras(
      `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/></wp:inline></w:drawing></w:r></w:p>${SECT}`,
    );
    expect(list[0]?.runs[0]?.content[0]).toEqual({
      kind: 'object',
      objectKind: 'drawing',
      width: 1440,
      height: 720,
    });
  });

  it('mc:AlternateContent 只取 mc:Choice —— 两个都收会画出两份图', () => {
    const { list } = paras(
      `<w:p><w:r><mc:AlternateContent>
        <mc:Choice><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/></wp:inline></w:drawing></mc:Choice>
        <mc:Fallback><w:pict/></mc:Fallback>
      </mc:AlternateContent></w:r></w:p>${SECT}`,
    );
    const content = list[0]?.runs[0]?.content ?? [];
    expect(content).toHaveLength(1);
    expect(content[0]).toMatchObject({ objectKind: 'drawing', width: 1440 });
  });
});

describe('透明容器一律压平', () => {
  it('超链接压成 run 上的标记，段落的子节点仍是扁平 run 列表', () => {
    const { list } = paras(
      `<w:p>
        <w:r><w:t>见</w:t></w:r>
        <w:hyperlink r:id="rId5" w:anchor="top"><w:r><w:t>这里</w:t></w:r></w:hyperlink>
      </w:p>${SECT}`,
    );
    const p = list[0] as Paragraph;
    expect(p.runs).toHaveLength(2);
    expect(p.runs[0]?.hyperlink).toBeUndefined();
    expect(p.runs[1]?.hyperlink).toEqual({ relId: 'rId5', anchor: 'top' });
    expect(paragraphText(p)).toBe('见这里');
  });

  it('w:ins / w:smartTag / w:fldSimple / w:sdt 里的 run 都要收上来', () => {
    const { list, sink } = paras(
      `<w:p>
        <w:ins><w:r><w:t>A</w:t></w:r></w:ins>
        <w:smartTag><w:r><w:t>B</w:t></w:r></w:smartTag>
        <w:fldSimple w:instr="PAGE"><w:r><w:t>C</w:t></w:r></w:fldSimple>
        <w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>D</w:t></w:r></w:sdtContent></w:sdt>
      </w:p>${SECT}`,
    );
    expect(paragraphText(list[0] as Paragraph)).toBe('ABCD');
    expect(sink.list()).toEqual([]);
  });

  it('修订容器压平成 run 上的标记：被删的字也进树（显不显示是级联的事，见 revisions.ts）', () => {
    const { list, sink } = paras(
      `<w:p><w:r><w:t>留</w:t></w:r>
         <w:del w:id="1" w:author="张三" w:date="2026-10-01T00:00:00Z"><w:r><w:delText>删</w:delText></w:r></w:del>
         <w:ins w:id="2" w:author="李四"><w:r><w:t>增</w:t></w:r></w:ins></w:p>${SECT}`,
    );
    const runs = (list[0] as Paragraph).runs;
    expect(runs.map((r) => r.content)).toEqual([
      [{ kind: 'text', text: '留' }],
      [{ kind: 'text', text: '删' }],
      [{ kind: 'text', text: '增' }],
    ]);
    // 段落文字答最终状态（paragraphText 的口径）
    expect(paragraphText(list[0] as Paragraph)).toBe('留增');
    expect(runs.map((r) => r.revision)).toEqual([
      undefined,
      { kind: 'delete', author: '张三', date: '2026-10-01T00:00:00Z' },
      { kind: 'insert', author: '李四' },
    ]);
    // 原来这里记一条 revision-deleted 的 info；字现在都在树上，不再有「没画出来」要留痕
    expect(sink.list()).toEqual([]);
  });

  it('移动记成带 move 的增删；嵌套时内层说了算；被删的域代码照样是域代码', () => {
    const { list } = paras(
      `<w:p><w:moveFrom w:id="1" w:author="甲"><w:r><w:delText>挪</w:delText></w:r></w:moveFrom>
         <w:moveTo w:id="2" w:author="甲"><w:r><w:t>挪</w:t></w:r></w:moveTo>
         <w:ins w:id="3" w:author="乙"><w:del w:id="4" w:author="丙"><w:r><w:delText>又删</w:delText></w:r></w:del></w:ins>
         <w:del w:id="5" w:author="丙"><w:r><w:fldChar w:fldCharType="begin"/></w:r>
           <w:r><w:delInstrText xml:space="preserve"> PAGE </w:delInstrText></w:r></w:del></w:p>${SECT}`,
    );
    const runs = (list[0] as Paragraph).runs;
    expect(runs.map((r) => r.revision)).toEqual([
      { kind: 'delete', move: true, author: '甲' },
      { kind: 'insert', move: true, author: '甲' },
      { kind: 'delete', author: '丙' },
      { kind: 'delete', author: '丙' },
      { kind: 'delete', author: '丙' },
    ]);
    expect(runs[4]?.content).toEqual([{ kind: 'fieldInstruction', text: ' PAGE ' }]);
  });

  it('块级 w:sdt 里的段落和表格照样进树', () => {
    const { body } = parse(
      `<w:sdt><w:sdtContent><w:p><w:r><w:t>X</w:t></w:r></w:p></w:sdtContent></w:sdt>${SECT}`,
    );
    expect([...walkParagraphs(body)].map((p) => paragraphText(p))).toEqual(['X']);
  });
});

describe('分节', () => {
  it('段落里的 sectPr 结束本节，且该段落属于本节；body 末尾的 sectPr 管最后一节', () => {
    const { body } = parse(
      `<w:p><w:pPr><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr></w:pPr>
         <w:r><w:t>一</w:t></w:r></w:p>
       <w:p><w:r><w:t>二</w:t></w:r></w:p>
       ${SECT}`,
    );
    expect(body.sections).toHaveLength(2);
    expect(body.sections[0]?.blocks).toHaveLength(1);
    expect(body.sections[0]?.props.page.orientation).toBe('landscape');
    // orient 只是声明，w:w 已经是横放后的宽度，不能再去转置
    expect(body.sections[0]?.props.page.width).toBe(16838);
    expect(blockText(body.sections[1]?.blocks[0])).toBe('二');
  });

  it('行网格照收 —— 公文「每页多少行」全靠它', () => {
    const { body } = parse(
      `<w:p/><w:sectPr><w:docGrid w:type="lines" w:linePitch="579" w:charSpace="1024"/></w:sectPr>`,
    );
    expect(body.sections[0]?.props.docGrid).toEqual({
      type: 'lines',
      linePitch: 579,
      charSpace: 1024,
    });
  });

  it('w:pgNumType 的起始页码与数字格式都收 —— 「前言罗马数字、正文阿拉伯数字」靠它', () => {
    const { body } = parse(`<w:p/><w:sectPr><w:pgNumType w:start="3" w:fmt="upperRoman"/></w:sectPr>`);
    expect(body.sections[0]?.props.pageNumStart).toBe(3);
    expect(body.sections[0]?.props.pageNumFormat).toBe('upperRoman');
    // 缺席就是缺席：写成 'decimal' 会让「域自己没写 \* 时用谁的格式」分不出「没配置」与「配了十进制」
    const bare = parse(`<w:p/><w:sectPr/>`);
    expect(bare.body.sections[0]?.props.pageNumFormat).toBeUndefined();
  });

  it('页眉页脚引用按类型收，没有 r:id 的丢掉', () => {
    const { body } = parse(
      `<w:p/><w:sectPr>
        <w:headerReference w:type="first" r:id="rId6"/>
        <w:headerReference w:type="default" r:id="rId7"/>
        <w:footerReference w:type="default"/>
      </w:sectPr>`,
    );
    expect(body.sections[0]?.props.headers).toEqual([
      { type: 'first', relId: 'rId6' },
      { type: 'default', relId: 'rId7' },
    ]);
    expect(body.sections[0]?.props.footers).toEqual([]);
  });

  it('完全没有 sectPr 时用兜底尺寸，并且要报出来', () => {
    const { body, sink } = parse(`<w:p><w:r><w:t>孤儿</w:t></w:r></w:p>`);
    expect(body.sections).toHaveLength(1);
    expect(body.sections[0]?.props.page.width).toBe(11906);
    expect(sink.list().map((d) => d.code)).toEqual(['missing-sectPr']);
  });
});

describe('表格', () => {
  const TBL = `<w:tbl>
      <w:tblPr/><w:tblGrid/>
      <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>乙</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>丙</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>${SECT}`;

  it('结构建全，单元格里的文字一个不少', () => {
    const { body, sink } = parse(TBL);
    const tbl = body.sections[0]?.blocks[0];
    expect(tbl?.kind).toBe('table');
    expect([...walkParagraphs(body)].map((p) => paragraphText(p))).toEqual(['甲', '乙', '丙']);
    expect(sink.list()).toEqual([]);
  });

  it('gridSpan 默认 1；<w:vMerge/> 不带 val 时是 continue，不是 restart', () => {
    const { body } = parse(TBL);
    const tbl = body.sections[0]?.blocks[0];
    if (tbl?.kind !== 'table') throw new Error('第一个块应当是表格');
    expect(tbl.rows[0]?.cells[0]).toMatchObject({ gridSpan: 2, vMerge: 'none' });
    expect(tbl.rows[1]?.cells[0]?.vMerge).toBe('restart');
    expect(tbl.rows[1]?.cells[1]?.vMerge).toBe('continue');
    expect(tbl.rows[1]?.cells[1]?.gridSpan).toBe(1);
  });
});

describe('未知元素', () => {
  it('报一条诊断后跳过，同名的只报一次 —— 否则一份文档能刷出上千条', () => {
    const { sink } = paras(`<w:p><w:zzz/><w:zzz/><w:r><w:qqq/></w:r></w:p>${SECT}`);
    const codes = sink.list();
    expect(codes.map((d) => d.path)).toEqual(['w:zzz', 'w:qqq']);
    expect(codes.every((d) => d.severity === 'warn' && d.part === 'document.xml')).toBe(true);
  });

  it('书签 / 拼写标记 / 批注范围不算未知 —— 每份文档都有一堆，报出来会淹掉真问题', () => {
    const { sink } = paras(
      `<w:bookmarkStart w:id="0" w:name="_Toc1"/>
       <w:p><w:proofErr w:type="spellStart"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="1"/></w:p>
       <w:bookmarkEnd w:id="0"/>${SECT}`,
    );
    expect(sink.list()).toEqual([]);
  });

  it('绝不采信 w:lastRenderedPageBreak —— 采信它等于让 Word 替我们排版', () => {
    const { list, sink } = paras(`<w:p><w:r><w:lastRenderedPageBreak/><w:t>x</w:t></w:r></w:p>${SECT}`);
    expect(list[0]?.runs[0]?.content).toEqual<RunContent[]>([{ kind: 'text', text: 'x' }]);
    expect(sink.list()).toEqual([]);
  });
});

describe('结构化克隆', () => {
  it('整棵树可结构化克隆 —— 这是过 Worker 边界的门票（原则 1.1）', () => {
    const { body } = parse(
      `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:hyperlink r:id="rId1"><w:r><w:t>a</w:t></w:r></w:hyperlink></w:p>
       <w:tbl><w:tr><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr></w:tbl>${SECT}`,
    );
    expect(structuredClone(body)).toEqual(body);
  });

  it('节点 id 在一棵树里唯一', () => {
    const { body } = parse(
      `<w:p><w:r><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r></w:p>
       <w:tbl><w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>${SECT}`,
    );
    const ids = body.sections.map((s) => s.id);
    for (const p of walkParagraphs(body)) ids.push(p.id, ...p.runs.map((r) => r.id));
    expect(new Set(ids).size).toBe(ids.length);
  });
});

describe('脚注 / 尾注', () => {
  it('引用成片段（带 id 与自定义标记），脚注内容开头的号也是片段；节上的编号规则收进 SectionProps', () => {
    const { body, sink } = parse(
      `<w:p><w:r><w:t>甲</w:t><w:footnoteReference w:id="3"/></w:r>` +
        `<w:r><w:footnoteReference w:customMarkFollows="1" w:id="4"/><w:t>*</w:t></w:r>` +
        `<w:r><w:endnoteReference w:id="1"/></w:r></w:p>` +
        `<w:sectPr><w:footnotePr><w:numFmt w:val="chineseCounting"/><w:numRestart w:val="eachSect"/></w:footnotePr></w:sectPr>`,
    );
    const runs = (body.sections[0]?.blocks[0] as Paragraph | undefined)?.runs ?? [];
    expect(runs.map((r) => r.content)).toEqual([
      [
        { kind: 'text', text: '甲' },
        { kind: 'noteReference', noteType: 'footnote', noteId: '3' },
      ],
      [
        { kind: 'noteReference', noteType: 'footnote', noteId: '4', customMark: true },
        { kind: 'text', text: '*' },
      ],
      [{ kind: 'noteReference', noteType: 'endnote', noteId: '1' }],
    ] satisfies RunContent[][]);
    expect(body.sections[0]?.props.footnotePr).toEqual({ numFmt: 'chineseCounting', numRestart: 'eachSect' });
    expect(body.sections[0]?.props.endnotePr).toBeUndefined();
    expect(sink.list()).toEqual([]);
  });

  it('footnotes.xml：分隔线那几条按 w:type 认出来，分隔线元素不报不认识；节点 id 带部件前缀且各条不撞', () => {
    const sink = createDiagnosticSink();
    const doc = parseXml(
      '<w:footnotes>' +
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
        '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>注一</w:t></w:r></w:p></w:footnote>' +
        '<w:footnote w:id="2"><w:p><w:r><w:t>注二</w:t></w:r></w:p></w:footnote>' +
        '</w:footnotes>',
      'footnotes.xml',
    );
    const notes = parseNotes(doc, sink, 'footnotes.xml', 'fn:');
    expect(notes.map((n) => [n.id, n.type, blockText(n.blocks[0])])).toEqual([
      ['-1', 'separator', ''],
      ['0', 'continuationSeparator', ''],
      ['1', 'normal', '注一'],
      ['2', 'normal', '注二'],
    ]);
    expect((notes[2]?.blocks[0] as Paragraph | undefined)?.runs[0]?.content).toEqual([
      { kind: 'noteMark', noteType: 'footnote' },
    ]);
    const ids = notes.flatMap((n) => n.blocks.map((b) => b.id));
    expect(new Set(ids).size).toBe(ids.length);
    expect(ids.every((id) => id.startsWith('fn:'))).toBe(true);
    expect(sink.list()).toEqual([]);
  });
});
