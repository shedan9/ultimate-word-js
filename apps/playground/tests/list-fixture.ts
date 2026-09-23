import { zipSync } from 'fflate';

const encoder = new TextEncoder();
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function rels(entries: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${RELS_NS}">${entries}</Relationships>`;
}

/** 三级十进制编号，每级缩进 420 twips、悬挂 420：降级后文字右移一格，浏览器里量得出来。 */
function numbering(): string {
  const levels = [0, 1, 2]
    .map(
      (l) => `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>
        <w:lvlText w:val="${Array.from({ length: l + 1 }, (_, i) => `%${i + 1}`).join('.')}."/>
        <w:pPr><w:ind w:left="${420 * (l + 1)}" w:hanging="420"/></w:pPr></w:lvl>`,
    )
    .join('');
  return `<w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="0">${levels}</w:abstractNum>
    <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
}

function paragraph(text: string, level?: number): string {
  const numPr =
    level === undefined
      ? ''
      : `<w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="1"/></w:numPr></w:pPr>`;
  return `<w:p>${numPr}${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p>`;
}

/** 列表编辑回归用：三个一级列表项 + 一段普通正文。 */
export function listDocx(): Uint8Array {
  const body = `<w:document xmlns:w="${W_NS}" xmlns:r="${OFFICE_REL}"><w:body>
    ${paragraph('第一项', 0)}${paragraph('第二项', 0)}${paragraph('第三项', 0)}${paragraph('普通正文')}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr></w:body></w:document>`;
  return zipSync({
    '[Content_Types].xml': encoder.encode(
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
      </Types>`,
    ),
    '_rels/.rels': encoder.encode(
      rels(`<Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="word/document.xml"/>`),
    ),
    'word/document.xml': encoder.encode(body),
    'word/_rels/document.xml.rels': encoder.encode(
      rels(`<Relationship Id="rId1" Type="${OFFICE_REL}/numbering" Target="numbering.xml"/>`),
    ),
    'word/numbering.xml': encoder.encode(numbering()),
  });
}
