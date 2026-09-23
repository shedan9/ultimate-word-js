import { readdirSync, readFileSync } from 'node:fs';
import { createDiagnosticSink } from '@uw/core';
import type { Body, DocPosition, Paragraph, TextTransaction } from '@uw/model';
import { createTextEditor, loadDocument, paragraphText, walkParagraphs } from '@uw/model';
import { OpcPackage, unzip } from '@uw/ooxml';
import { zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { serializeDocx } from './docx.ts';

const FIXTURES = new URL('../../../apps/fidelity/fixtures/', import.meta.url);
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const encoder = new TextEncoder();
const decoder = new TextDecoder();

function fixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(new URL(name, FIXTURES)));
}

/** 最小 docx：只有正文（外加可选的 numbering.xml） */
function docx(body: string, numbering?: string): Uint8Array {
  const rel = (id: string, type: string, target: string) =>
    `<Relationship Id="${id}" Type="${R_NS}/${type}" Target="${target}"/>`;
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': encoder.encode(
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        (numbering === undefined
          ? ''
          : `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`) +
        `</Types>`,
    ),
    '_rels/.rels': encoder.encode(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId1', 'officeDocument', 'word/document.xml')}</Relationships>`,
    ),
    'word/document.xml': encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`,
    ),
    'word/_rels/document.xml.rels': encoder.encode(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId9', 'hyperlink', 'https://example.com" TargetMode="External')}${numbering === undefined ? '' : rel('rId1', 'numbering', 'numbering.xml')}</Relationships>`,
    ),
  };
  if (numbering !== undefined)
    files['word/numbering.xml'] = encoder.encode(`<w:numbering xmlns:w="${W_NS}">${numbering}</w:numbering>`);
  return zipSync(files);
}

function load(bytes: Uint8Array) {
  const pkg = OpcPackage.open(bytes);
  return { pkg, loaded: loadDocument(pkg, createDiagnosticSink()) };
}

function paragraphs(body: Body): Paragraph[] {
  return [...walkParagraphs(body)];
}

/** 结构比较用：id 是按解析顺序现编的，重新加载后必然不同；删空的文字片段不写进文件 */
function shape(body: Body): unknown {
  return JSON.parse(
    JSON.stringify(body.sections, (key, value) => {
      if (key === 'id') return undefined;
      if (key === 'content' && Array.isArray(value))
        return value.filter((c) => c.kind !== 'text' || c.text !== '');
      return value;
    }),
  );
}

function roundTrip(bytes: Uint8Array, edit: (t: TextTransaction, body: Body) => void) {
  const { pkg, loaded } = load(bytes);
  const editor = createTextEditor(loaded.body);
  editor.tx((t) => {
    edit(t, editor.body);
    return undefined;
  });
  const out = serializeDocx(pkg, editor.body);
  const again = load(out);
  return { edited: editor.body, out, again: again.loaded, pkg: again.pkg };
}

/** 第一个带文字的段落的某个字缝 */
function textPosition(body: Body, nth = 0, offset = 1): DocPosition | undefined {
  let seen = 0;
  for (const p of paragraphs(body)) {
    for (const r of p.runs) {
      const ci = r.content.findIndex((c) => c.kind === 'text' && c.text.length > offset);
      if (ci < 0) continue;
      if (seen++ === nth) return { nodeId: r.id, contentIndex: ci, offset };
    }
  }
  return undefined;
}

const docxFixtures = readdirSync(FIXTURES).filter((f) => f.endsWith('.docx'));

describe('不编辑：逐字节照搬', () => {
  it.each(docxFixtures)('%s 的每个部件与原包相同', (name) => {
    const bytes = fixture(name);
    const { pkg, loaded } = load(bytes);
    const before = unzip(bytes);
    const after = unzip(serializeDocx(pkg, createTextEditor(loaded.body).body));
    expect([...after.keys()]).toEqual([...before.keys()]);
    for (const [k, v] of before) expect(after.get(k), k).toEqual(v);
  });
});

describe('编辑后重新加载，正文结构与编辑结果一致', () => {
  it.each(docxFixtures)('%s：输入 + 拆段 + 加粗 + 对齐', (name) => {
    const { edited, again } = roundTrip(fixture(name), (t, body) => {
      const at = textPosition(body);
      if (at === undefined) return;
      const caret = t.insertText(at, '插入<&>');
      const bold = t.setRunProps({ start: at, end: caret }, { bold: true, size: 420 });
      const split = t.splitParagraph(bold.end);
      t.insertInline(split, 'tab');
      t.setParagraphProps(
        { start: split, end: split },
        { justification: 'center', indent: { firstLineChars: 200 } },
      );
    });
    expect(shape(again.body)).toEqual(shape(edited));
  });

  it('合段：后一段的 run 搬进前一段，超链接跟着过去', () => {
    const bytes = docx(
      '<w:p><w:r><w:t>甲段</w:t></w:r></w:p>' +
        '<w:p><w:hyperlink r:id="rId9"><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>链接</w:t></w:r></w:hyperlink><w:r><w:t>尾巴</w:t></w:r></w:p>',
    );
    const { edited, again, out } = roundTrip(bytes, (t, body) => {
      t.joinParagraph((paragraphs(body)[0] as Paragraph).id);
    });
    expect(shape(again.body)).toEqual(shape(edited));
    expect(paragraphs(again.body).map((p) => paragraphText(p))).toEqual(['甲段链接尾巴']);
    expect(paragraphs(again.body)[0]?.runs[1]?.hyperlink).toEqual({ relId: 'rId9' });
    expect(decoder.decode(unzip(out).get('word/document.xml'))).toContain('<w:hyperlink r:id="rId9">');
  });
});

describe('空段落', () => {
  it('在空段落里输入：新 run 带着段落标记的格式写出来，书签留着', () => {
    const bytes = docx(
      '<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:bookmarkStart w:id="0" w:name="空"/><w:bookmarkEnd w:id="0"/></w:p>',
    );
    const { edited, again, out } = roundTrip(bytes, (t, body) => {
      t.insertText({ nodeId: (paragraphs(body)[0] as Paragraph).id, contentIndex: 0, offset: 0 }, '有字了');
    });
    expect(shape(again.body)).toEqual(shape(edited));
    const xml = decoder.decode(unzip(out).get('word/document.xml'));
    expect(xml).toContain('<w:bookmarkStart w:id="0" w:name="空"/>');
    expect(xml).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t>有字了</w:t></w:r>');
  });
});

describe('模型不认识的 XML 留在文件里', () => {
  const PRESERVED =
    '<w:p w14:paraId="1A2B3C4D" w:rsidR="00AB12CD">' +
    '<w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr><w:jc w:val="start"/></w:pPr>' +
    '<w:bookmarkStart w:id="0" w:name="书签"/>' +
    '<w:r><w:rPr><w:highlight w:val="yellow"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr><w:t>高亮文字</w:t></w:r>' +
    '<w:proofErr w:type="spellStart"/>' +
    '<w:del w:id="1" w:author="甲"><w:r><w:delText>删掉的</w:delText></w:r></w:del>' +
    '<w:r><w:footnoteReference w:id="2"/></w:r>' +
    '<w:bookmarkEnd w:id="0"/>' +
    '</w:p>';

  it('改了文字的段落：书签、修订、边框、高亮、脚注引用、paraId 全在', () => {
    const { out, again } = roundTrip(docx(PRESERVED), (t, body) => {
      t.insertText(textPosition(body) as DocPosition, '新');
    });
    const xml = decoder.decode(unzip(out).get('word/document.xml'));
    for (const piece of [
      'w14:paraId="1A2B3C4D"',
      '<w:pBdr>',
      // 对齐没改：原来的新写法 start 不被改写成 left
      '<w:jc w:val="start"/>',
      '<w:bookmarkStart w:id="0" w:name="书签"/>',
      '<w:bookmarkEnd w:id="0"/>',
      '<w:highlight w:val="yellow"/>',
      '<w:lang w:val="en-US" w:eastAsia="zh-CN"/>',
      '<w:proofErr w:type="spellStart"/>',
      '<w:delText>删掉的</w:delText>',
      '<w:footnoteReference w:id="2"/>',
    ])
      expect(xml).toContain(piece);
    expect(paragraphText(paragraphs(again.body)[0] as Paragraph)).toBe('高新亮文字');
  });

  it('改格式只动那一个属性：w:lang 的 w:val 留着，新元素按 schema 顺序插', () => {
    const { out } = roundTrip(docx(PRESERVED), (t, body) => {
      const at = textPosition(body, 0, 0) as DocPosition;
      t.setRunProps({ start: at, end: { ...at, offset: 4 } }, { bold: true, langEastAsia: 'ja-JP' });
    });
    const xml = decoder.decode(unzip(out).get('word/document.xml'));
    // w:b 排在 w:highlight 之前（EG_RPrBase 的顺序），w:lang 只换了 eastAsia
    expect(xml).toContain(
      '<w:rPr><w:b/><w:highlight w:val="yellow"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr>',
    );
  });

  it('拆出来的新段落沿用原段落的边框，但不抄 paraId', () => {
    const { out, again } = roundTrip(docx(PRESERVED), (t, body) => {
      t.splitParagraph(textPosition(body, 0, 2) as DocPosition);
    });
    const xml = decoder.decode(unzip(out).get('word/document.xml'));
    expect(xml.match(/<w:pBdr>/g)).toHaveLength(2);
    expect(xml.match(/w14:paraId=/g)).toHaveLength(1);
    // 新段落的 run 以拆点所在的 run 为模板：高亮跟过去
    expect(xml.match(/<w:highlight w:val="yellow"\/>/g)).toHaveLength(2);
    expect(paragraphs(again.body).map((p) => paragraphText(p))).toEqual(['高亮', '文字']);
  });

  it('块级内容控件：外壳留着，拆出来的段落留在控件里', () => {
    const bytes = docx(
      '<w:sdt><w:sdtPr><w:alias w:val="坑位"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>控件里</w:t></w:r></w:p></w:sdtContent></w:sdt>' +
        '<w:p><w:r><w:t>控件外</w:t></w:r></w:p>',
    );
    const { out, again } = roundTrip(bytes, (t, body) => {
      t.splitParagraph(textPosition(body, 0, 1) as DocPosition);
    });
    const xml = decoder.decode(unzip(out).get('word/document.xml'));
    expect(xml).toMatch(/<w:sdtContent><w:p>.*控.*<\/w:p><w:p>.*件里.*<\/w:p><\/w:sdtContent><\/w:sdt><w:p>/);
    expect(xml).toContain('<w:alias w:val="坑位"/>');
    expect(paragraphs(again.body)).toHaveLength(3);
  });
});

describe('分节', () => {
  const TWO_SECTIONS =
    '<w:p><w:r><w:t>第一节末段</w:t></w:r><w:pPr><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr></w:pPr></w:p>' +
    '<w:p><w:r><w:t>第二节</w:t></w:r></w:p>';

  it('拆第一节的末段：分节符搬到新的末段上，节数与每节的块数不变形', () => {
    const { edited, again, out } = roundTrip(docx(TWO_SECTIONS), (t, body) => {
      t.splitParagraph(textPosition(body, 0, 2) as DocPosition);
    });
    expect(again.body.sections.map((s) => s.blocks.length)).toEqual([2, 1]);
    expect(again.body.sections[0]?.props.page.orientation).toBe('landscape');
    expect(shape(again.body)).toEqual(shape(edited));
    const xml = decoder.decode(unzip(out).get('word/document.xml'));
    expect(xml.match(/<w:sectPr>/g)).toHaveLength(2);
  });
});

describe('编号定义', () => {
  it('新建列表追加到已有 numbering.xml：abstractNum 在所有 num 之前', () => {
    const bytes = docx(
      '<w:p><w:r><w:t>一项</w:t></w:r></w:p>',
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
    );
    const { edited, again, out } = roundTrip(bytes, (t, body) => {
      const numId = t.addList('bullet');
      const at = textPosition(body, 0, 0) as DocPosition;
      t.setParagraphProps({ start: at, end: at }, { numbering: { numId, level: 0 } });
    });
    const xml = decoder.decode(unzip(out).get('word/numbering.xml'));
    expect(xml.lastIndexOf('<w:abstractNum ')).toBeLessThan(xml.indexOf('<w:num '));
    expect(again.body.numbering).toEqual(edited.numbering);
    expect(again.resolved.sections[0]?.blocks[0]).toMatchObject({
      props: { numbering: { label: { text: '●' } } },
    });
  });

  it('原包没有 numbering.xml：新建部件并登记内容类型与关系', () => {
    const { again, out } = roundTrip(docx('<w:p><w:r><w:t>一项</w:t></w:r></w:p>'), (t, body) => {
      const numId = t.addList('decimal');
      const at = textPosition(body, 0, 0) as DocPosition;
      t.setParagraphProps({ start: at, end: at }, { numbering: { numId, level: 0 } });
    });
    const entries = unzip(out);
    expect(entries.has('word/numbering.xml')).toBe(true);
    expect(decoder.decode(entries.get('[Content_Types].xml'))).toContain('PartName="/word/numbering.xml"');
    // rId9 已被超链接占了，新关系另取一个
    expect(decoder.decode(entries.get('word/_rels/document.xml.rels'))).toMatch(
      /Id="rId1" Type="[^"]+\/numbering"/,
    );
    expect(again.resolved.sections[0]?.blocks[0]).toMatchObject({
      props: { numbering: { label: { text: '1.' } } },
    });
  });

  it('撤销掉的新建列表不留定义', () => {
    const { pkg, loaded } = load(docx('<w:p><w:r><w:t>一项</w:t></w:r></w:p>'));
    const editor = createTextEditor(loaded.body);
    editor.tx((t) => {
      const numId = t.addList('decimal');
      const at = textPosition(editor.body, 0, 0) as DocPosition;
      t.setParagraphProps({ start: at, end: at }, { numbering: { numId, level: 0 } });
      return undefined;
    });
    editor.undo();
    expect(unzip(serializeDocx(pkg, editor.body)).has('word/numbering.xml')).toBe(false);
  });
});

describe('图片', () => {
  it('拆开含图的 run 之后，图形元素原样写回', () => {
    const bytes = fixture('spike-image-01.docx');
    const { again, out } = roundTrip(bytes, (t, body) => {
      // 全文加粗：每个含图的 run 都被改写，图只能从原元素池里找回
      const ps = paragraphs(body);
      const first = ps[0]?.runs[0];
      const lastP = ps.at(-1) as Paragraph;
      const last = lastP.runs.at(-1);
      if (first === undefined || last === undefined) throw new Error('样本没有 run');
      const ci = last.content.length - 1;
      const tail = last.content[ci];
      const offset = tail?.kind === 'text' ? tail.text.length : 1;
      t.setRunProps(
        {
          start: { nodeId: first.id, contentIndex: 0, offset: 0 },
          end: { nodeId: last.id, contentIndex: ci, offset },
        },
        { bold: true },
      );
    });
    const count = (s: string) => (s.match(/<w:drawing>/g) ?? []).length;
    expect(count(decoder.decode(unzip(out).get('word/document.xml')))).toBe(
      count(decoder.decode(unzip(bytes).get('word/document.xml'))),
    );
    expect(Object.keys(again.images).length).toBeGreaterThan(0);
  });
});
