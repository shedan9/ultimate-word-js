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

const SECT_PR = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>`;

/** 最小的 docx：一个正文部件 + 编号部件；`extraRels` 追加进正文的关系表（超链接用）。 */
function docxOf(blocks: string, extraRels = ''): Uint8Array {
  const body = `<w:document xmlns:w="${W_NS}" xmlns:r="${OFFICE_REL}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${blocks}${SECT_PR}</w:body></w:document>`;
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
      rels(`<Relationship Id="rId1" Type="${OFFICE_REL}/numbering" Target="numbering.xml"/>${extraRels}`),
    ),
    'word/numbering.xml': encoder.encode(numbering()),
  });
}

/** 文字模板状态回归：普通输入必须同步清占位标记与绑定。 */
export function textStateDocx(type: 'text' | 'richText'): Uint8Array {
  return docxOf(
    `<w:p><w:sdt><w:sdtPr><w:tag w:val="field"/><w:showingPlcHdr/><w:dataBinding w:xpath="/field"/><w:${type}/></w:sdtPr><w:sdtContent><w:r><w:t>占位文字</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
  );
}

/** 列表编辑回归用：三个一级列表项 + 一段普通正文。 */
export function listDocx(): Uint8Array {
  return docxOf(
    `${paragraph('第一项', 0)}${paragraph('第二项', 0)}${paragraph('第三项', 0)}${paragraph('普通正文')}`,
  );
}

/** 列表模板：按选项值填写后，点击、撤销与回写应同步保存值。 */
export function choiceControlDocx(type: 'dropDownList' | 'comboBox' = 'dropDownList'): Uint8Array {
  return docxOf(
    `<w:p><w:sdt><w:sdtPr><w:tag w:val="level"/><w:${type} w:lastValue="2">` +
      '<w:listItem w:displayText="紧急" w:value="1"/><w:listItem w:displayText="一般" w:value="2"/>' +
      `</w:${type}></w:sdtPr><w:sdtContent><w:r><w:t>一般</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
  );
}

/** 填写面板回归：同名字段、重复显示 / 保存值冲突，以及各类可写与不可写状态。 */
export function templateControlsDocx(): Uint8Array {
  const control = (tag: string, text: string, pr: string) =>
    `<w:p><w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${pr}</w:sdtPr><w:sdtContent><w:r><w:t>${text}</w:t></w:r></w:sdtContent></w:sdt></w:p>`;
  const items =
    '<w:listItem w:displayText="一般" w:value="a"/><w:listItem w:displayText="一般" w:value="b"/><w:listItem w:displayText="b" w:value="c"/>';
  return docxOf(
    paragraph('项目申请表') +
      control('紧急程度', '一般', `<w:dropDownList w:lastValue="a">${items}</w:dropDownList>`) +
      control('紧急程度', '一般', `<w:dropDownList w:lastValue="a">${items}</w:dropDownList>`) +
      control('办理意见', '待确认', '<w:comboBox/>') +
      control(
        '签发日期',
        '2026年10月16日',
        '<w:date w:fullDate="2026-10-16T00:00:00Z"><w:dateFormat w:val="yyyy年M月d日"/></w:date>',
      ).replace(
        '<w:t>2026年10月16日</w:t>',
        '<w:rPr><w:b/></w:rPr><w:t>2026年</w:t></w:r><w:r><w:t>10月16日</w:t>',
      ) +
      control('同意申请', '☐', '<w14:checkbox/>') +
      `<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>${control('审批结果', '已归档', '<w:comboBox/>')}</w:sdtContent></w:sdt>` +
      control('空选项', '请选择', '<w:showingPlcHdr/><w:dropDownList/>') +
      control('缺少日期格式', '某日', '<w:date/>') +
      control('申请说明', '文字字段请在编辑模式下填写', '<w:text/>'),
  );
}

/** 日期模板：点击事件应读到填写后的元数据，显示文字与撤销 / 回写同进同退。 */
export function dateControlDocx(): Uint8Array {
  return docxOf(
    '<w:p><w:sdt><w:sdtPr><w:tag w:val="issued"/><w:showingPlcHdr/>' +
      '<w:date><w:dateFormat w:val="yyyy年M月d日"/><w:lid w:val="zh-CN"/><w:calendar w:val="gregorian"/></w:date>' +
      '</w:sdtPr><w:sdtContent><w:r><w:t>某日</w:t></w:r></w:sdtContent></w:sdt></w:p>',
  );
}

/** 表格 Tab 导航回归用：一段正文 + 2×2 表格（第二行首格带列表段落，Tab 要跳格而不是降级）。 */
export function tableDocx(): Uint8Array {
  const cell = (inner: string) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
  const border = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`;
  return docxOf(
    `${paragraph('表前正文')}<w:tbl><w:tblPr><w:tblBorders>${[
      'top',
      'left',
      'bottom',
      'right',
      'insideH',
      'insideV',
    ]
      .map(border)
      .join('')}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>
      <w:tr>${cell(paragraph('格甲'))}${cell(paragraph('格乙'))}</w:tr>
      <w:tr>${cell(paragraph('格丙', 0))}${cell(paragraph('格丁'))}</w:tr></w:tbl>${paragraph('表后正文')}`,
  );
}

/**
 * 事件回归用：外链（地址在关系表里、`TargetMode="External"`）+ 书签跳转（只有 `w:anchor`）+ 普通文字。
 * 两种超链接走的是 `w:hyperlink` 容器那条路，HYPERLINK 域那条由 model 的单测兜着。
 */
export function linkDocx(): Uint8Array {
  return docxOf(
    `<w:p><w:r><w:t>请访问</w:t></w:r><w:hyperlink r:id="rId9"><w:r><w:t>官方网站</w:t></w:r></w:hyperlink></w:p>
      <w:p><w:hyperlink w:anchor="附件"><w:r><w:t>跳到附件</w:t></w:r></w:hyperlink></w:p>${paragraph('普通正文')}`,
    `<Relationship Id="rId9" Type="${OFFICE_REL}/hyperlink" Target="https://example.com/" TargetMode="External"/>`,
  );
}

/** 坑位点击用：行内 / 块级、嵌套内容锁、空段、表格与控件里的超链接。 */
export function controlDocx(): Uint8Array {
  const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
  const control = (tag: string, content: string, props = '<w:text/>') =>
    `<w:sdt><w:sdtPr>${tag ? `<w:tag w:val="${tag}"/>` : ''}${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  return docxOf(
    `<w:p>${control('applicant', run('申请人'))}${run('普通文字')}</w:p>
    ${control('outer', `<w:p>${run('外层文字')}${control('inner', run('内层文字'))}</w:p>`, '<w:lock w:val="contentLocked"/>')}
    <w:p>${control('agree', run('☐'), '<w14:checkbox/>')}</w:p>
    ${control('empty', '<w:p/>')}
    ${control('table', '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>格内文字</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')}
    <w:p>${control('', run('无名控件'))}</w:p>
    <w:p>${control('linked', `<w:hyperlink w:anchor="附件">${run('控件内链接')}</w:hyperlink>`, '')}</w:p>
    <w:p>${control('shell', run('外壳锁'), '<w:text/><w:lock w:val="sdtLocked"/>')}</w:p>`,
  );
}

/** 块级合段与外层填充会移除外壳；外壳内部仍须允许普通输入。 */
export function shellControlDocx(): Uint8Array {
  const control = (tag: string, content: string, locked = true) =>
    `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${locked ? '<w:lock w:val="sdtLocked"/>' : ''}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  return docxOf(
    `${paragraph('前段')}${control('blockShell', paragraph('块内文字'))}` +
      control(
        'outerShell',
        `<w:p><w:r><w:t>外层前缀</w:t></w:r>${control('inlineShell', '<w:r><w:t>行内文字</w:t></w:r>')}</w:p>`,
        false,
      ),
  );
}
