/**
 * 批注（`doc.comments()`）：谁批的、批了什么、罩着哪些字，编辑之后与回写之后还罩着同一段字。
 *
 * 判据是接线 —— 批注不排版、没有真值，只看三件事：① 范围交叠 / 跨段 / 点批注都认得出；
 * ② 回复与「已解决」从 `commentsExtended.xml` 按 paraId 接上；③ 标记跟着 run 走，
 * 编辑后的范围、回写后再加载的范围都对得上。
 */
import { unzip, zip } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { UwDocument } from './index.ts';
import { UltimateWord } from './index.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';

const start = (id: string) => `<w:commentRangeStart w:id="${id}"/>`;
const end = (id: string) => `<w:commentRangeEnd w:id="${id}"/>`;
const ref = (id: string) =>
  `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const comment = (id: string, author: string, paraId: string, ...paras: string[]) =>
  `<w:comment w:id="${id}" w:author="${author}" w:initials="${author.slice(0, 1)}" w:date="2026-10-0${Number(id) + 1}T09:00:00Z">` +
  paras
    .map(
      (text, i) =>
        `<w:p${i === paras.length - 1 ? ` w14:paraId="${paraId}"` : ''}>` +
        `${i === 0 ? '<w:r><w:annotationRef/></w:r>' : ''}${run(text)}</w:p>`,
    )
    .join('') +
  '</w:comment>';

/**
 * 0 罩「甲乙丙」，回复 3 与它同一个范围；1 从「丙」罩到下一段的「丁」（与 0 交叠、跨段）；
 * 2 是点批注（范围是空的，只有引用标记）。1 已解决；3 回复 0
 */
function commentDocx(plain = false): Uint8Array {
  const enc = new TextEncoder();
  const body = plain
    ? `<w:p>${run('甲乙丙')}</w:p><w:p>${run('丁戊己')}</w:p>`
    : `<w:p>${start('0')}${start('3')}${run('甲乙')}${start('1')}${run('丙')}${end('0')}${end('3')}${ref('0')}${ref('3')}</w:p>` +
      `<w:p>${run('丁')}${end('1')}${ref('1')}${run('戊')}${start('2')}${end('2')}${ref('2')}${run('己')}</w:p>`;
  const ct = (part: string, type: string) => `<Override PartName="/word/${part}" ContentType="${type}"/>`;
  const wml = (t: string) => `application/vnd.openxmlformats-officedocument.wordprocessingml.${t}+xml`;
  const rel = (id: string, type: string, target: string) =>
    `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
  const files = new Map<string, Uint8Array>([
    [
      '[Content_Types].xml',
      enc.encode(
        '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          ct('document.xml', wml('document.main')) +
          ct('comments.xml', wml('comments')) +
          ct('commentsExtended.xml', wml('commentsExtended')) +
          '</Types>',
      ),
    ],
    [
      '_rels/.rels',
      enc.encode(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId1', `${R}/officeDocument`, 'word/document.xml')}</Relationships>`,
      ),
    ],
    [
      'word/_rels/document.xml.rels',
      enc.encode(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          rel('rId1', `${R}/comments`, 'comments.xml') +
          rel(
            'rId2',
            'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
            'commentsExtended.xml',
          ) +
          '</Relationships>',
      ),
    ],
    [
      'word/comments.xml',
      enc.encode(
        `<w:comments xmlns:w="${W}" xmlns:w14="${W14}">` +
          comment('0', '张三', '00000A00', '这里要改') +
          comment('1', '李四', '00000B00', '第一段', '第二段') +
          comment('2', '王五', '00000C00', '补一句') +
          comment('3', '李四', '00000D00', '同意') +
          '</w:comments>',
      ),
    ],
    [
      'word/commentsExtended.xml',
      enc.encode(
        `<w15:commentsEx xmlns:w15="${W15}">` +
          '<w15:commentEx w15:paraId="00000A00" w15:done="0"/>' +
          '<w15:commentEx w15:paraId="00000B00" w15:done="1"/>' +
          '<w15:commentEx w15:paraId="00000D00" w15:paraIdParent="00000A00" w15:done="0"/>' +
          '</w15:commentsEx>',
      ),
    ],
    [
      'word/document.xml',
      enc.encode(
        `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
      ),
    ],
  ]);
  if (plain) {
    // 原来没有批注的文档：批注部件、关系、内容类型都没有，新建批注要把它们补齐
    files.delete('word/comments.xml');
    files.delete('word/commentsExtended.xml');
    files.set(
      'word/_rels/document.xml.rels',
      enc.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'),
    );
    const types = new TextDecoder().decode(files.get('[Content_Types].xml'));
    files.set(
      '[Content_Types].xml',
      enc.encode(types.replace(/<Override PartName="\/word\/comments[^>]*>/g, '')),
    );
  }
  return zip(files);
}

describe('doc.comments', () => {
  it('作者 / 时间 / 正文 / 回复 / 已解决，按锚点的文档序；交叠、跨段、点批注都认得出', async () => {
    const d = await UltimateWord.load(commentDocx());
    const list = d.comments();
    expect(list.map((c) => c.id)).toEqual(['0', '3', '1', '2']);
    expect(list.map(({ range: _, ...rest }) => rest)).toEqual([
      {
        id: '0',
        author: '张三',
        initials: '张',
        date: '2026-10-01T09:00:00Z',
        done: false,
        text: '这里要改',
        quote: '甲乙丙',
      },
      {
        id: '3',
        author: '李四',
        initials: '李',
        date: '2026-10-04T09:00:00Z',
        parent: '0',
        done: false,
        text: '同意',
        quote: '甲乙丙',
      },
      {
        id: '1',
        author: '李四',
        initials: '李',
        date: '2026-10-02T09:00:00Z',
        done: true,
        text: '第一段\n第二段',
        quote: '丙\n丁',
      },
      {
        id: '2',
        author: '王五',
        initials: '王',
        date: '2026-10-03T09:00:00Z',
        done: false,
        text: '补一句',
        quote: '',
      },
    ]);
    // 点批注是折叠的范围，落在「戊」的结尾 —— 不落在引用标记那个空 run 上（布局里没有它，锚不住）；
    // 有范围的批注两端也都落在有字的 run 上
    const wu = d.find('戊')[0];
    expect(list[3]?.range).toEqual({ start: wu?.end, end: wu?.end });
    expect(list[0]?.range).toEqual({ start: d.find('甲乙')[0]?.start, end: d.find('丙')[0]?.end });
    // 批注不进正文：排出来的字里没有批注内容
    expect(d.find('这里要改')).toEqual([]);
  });

  it('编辑后范围跟着字走：范围里插字跟着变长，撤销复原', async () => {
    const d = await UltimateWord.load(commentDocx());
    const hit = d.find('乙')[0];
    if (hit === undefined) throw new Error('找不到「乙」');
    d.tx((t) => {
      t.insertText(hit.start, '和');
    });
    expect(d.comments().find((c) => c.id === '0')?.quote).toBe('甲和乙丙');
    d.undo();
    expect(d.comments().find((c) => c.id === '0')?.quote).toBe('甲乙丙');
  });

  it('回写：不编辑时 comments.xml 逐字节照搬；编辑过的段落里批注界桩还在，再加载范围一致', async () => {
    const src = commentDocx();
    const d = await UltimateWord.load(src);
    const hit = d.find('乙')[0];
    if (hit === undefined) throw new Error('找不到「乙」');
    d.tx((t) => {
      t.insertText(hit.start, '和');
    });
    const out = new Uint8Array(await (await d.toDocx()).arrayBuffer());
    const parts = unzip(out);
    const original = unzip(src);
    expect(parts.get('word/comments.xml')).toEqual(original.get('word/comments.xml'));
    expect(parts.get('word/commentsExtended.xml')).toEqual(original.get('word/commentsExtended.xml'));
    const again = await UltimateWord.load(out);
    const pick = (x: UwDocument) => x.comments().map((c) => [c.id, c.quote, c.parent, c.done]);
    expect(pick(again)).toEqual(pick(d));
    expect(pick(again)[0]).toEqual(['0', '甲和乙丙', undefined, false]);
  });
});

const text = (parts: Map<string, Uint8Array>, name: string) => new TextDecoder().decode(parts.get(name));
const bytesOf = async (d: UwDocument) => new Uint8Array(await (await d.toDocx()).arrayBuffer());
const pick = (x: UwDocument) => x.comments().map((c) => [c.id, c.quote, c.parent, c.done, c.text, c.author]);

describe('批注的编辑', () => {
  it('新建 + 回复 + 解决：一个命令一个撤销单元，回写再加载一模一样；格式照文件里已有的批注抄', async () => {
    const d = await UltimateWord.load(commentDocx(), { author: '赵六' });
    const hit = d.find('戊')[0];
    if (hit === undefined) throw new Error('找不到「戊」');
    const id = d.addComment(hit, '这个字', { date: '2026-10-12T10:00:00Z' });
    expect(id).toBe('4');
    const reply = d.replyComment(id, '收到', { author: '钱七', date: '2026-10-12T11:00:00Z' });
    d.resolveComment(reply);
    const added = d.comments().filter((c) => Number(c.id) >= 4);
    expect(added.map(({ range: _, ...rest }) => rest)).toEqual([
      {
        id: '4',
        author: '赵六',
        initials: '赵',
        date: '2026-10-12T10:00:00Z',
        done: true,
        text: '这个字',
        quote: '戊',
      },
      {
        id: '5',
        author: '钱七',
        initials: '钱',
        date: '2026-10-12T11:00:00Z',
        parent: '4',
        done: true,
        text: '收到',
        quote: '戊',
      },
    ]);

    const out = await bytesOf(d);
    const again = await UltimateWord.load(out);
    expect(pick(again)).toEqual(pick(d));
    const parts = unzip(out);
    const doc = text(parts, 'word/document.xml');
    // Word 的写法：界桩夹着字、引用紧跟终点；两条批注罩着同一个字
    expect(doc).toContain(
      '<w:commentRangeStart w:id="4"/><w:commentRangeStart w:id="5"/><w:r><w:t>戊</w:t></w:r>' +
        '<w:commentRangeEnd w:id="4"/><w:commentRangeEnd w:id="5"/>' +
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="4"/></w:r>',
    );
    const comments = text(parts, 'word/comments.xml');
    expect(comments).toContain(
      '<w:comment w:id="4" w:author="赵六" w:date="2026-10-12T10:00:00Z" w:initials="赵">',
    );
    expect(comments).toContain('<w:r><w:annotationRef/></w:r><w:r><w:t>这个字</w:t></w:r>');
    // 原有的四条一个字节都没动
    for (const c of ['0', '1', '2', '3'])
      expect(comments).toContain(
        text(unzip(commentDocx()), 'word/comments.xml').match(
          new RegExp(`<w:comment w:id="${c}".*?</w:comment>`),
        )?.[0],
      );
    const ex = text(parts, 'word/commentsExtended.xml');
    expect(ex).toMatch(
      /<w15:commentEx w15:paraId="([0-9A-F]{8})" w15:done="1"\/><w15:commentEx w15:paraId="[0-9A-F]{8}" w15:paraIdParent="\1" w15:done="1"\/>/,
    );

    // 撤销三次回到原样：回写的批注部件与原文逐字节相同
    d.undo();
    d.undo();
    d.undo();
    const back = unzip(await bytesOf(d));
    expect(back.get('word/comments.xml')).toEqual(unzip(commentDocx()).get('word/comments.xml'));
    expect(back.get('word/document.xml')).toEqual(unzip(commentDocx()).get('word/document.xml'));
  });

  it('删除：删顶层连回复一起删，界桩、引用、扩展条目都摘掉；别的批注照旧', async () => {
    const d = await UltimateWord.load(commentDocx());
    d.deleteComment('0');
    expect(d.comments().map((c) => c.id)).toEqual(['1', '2']);
    const parts = unzip(await bytesOf(d));
    const doc = text(parts, 'word/document.xml');
    expect(doc).not.toMatch(/w:id="[03]"/);
    expect(doc).toContain('<w:commentRangeStart w:id="1"/>');
    const comments = text(parts, 'word/comments.xml');
    expect(comments).not.toMatch(/<w:comment w:id="[03]"/);
    const ex = text(parts, 'word/commentsExtended.xml');
    expect(ex).not.toContain('00000A00');
    expect(ex).not.toContain('00000D00');
    expect(ex).toContain('<w15:commentEx w15:paraId="00000B00" w15:done="1"/>');
    const again = await UltimateWord.load(await bytesOf(d));
    expect(pick(again)).toEqual(pick(d));
  });

  it('改正文：段落重写、末段的 paraId 留着（回复与解决状态还认得它）', async () => {
    const d = await UltimateWord.load(commentDocx());
    d.editComment('1', '改过了\n两段');
    const parts = unzip(await bytesOf(d));
    const comments = text(parts, 'word/comments.xml');
    expect(comments).toContain(
      '<w:p><w:r><w:annotationRef/></w:r><w:r><w:t>改过了</w:t></w:r></w:p><w:p w14:paraId="00000B00"><w:r><w:t>两段</w:t></w:r></w:p>',
    );
    // 扩展部件没变，原样照搬
    expect(parts.get('word/commentsExtended.xml')).toEqual(
      unzip(commentDocx()).get('word/commentsExtended.xml'),
    );
    const again = await UltimateWord.load(await bytesOf(d));
    expect(again.comments().find((c) => c.id === '1')).toMatchObject({ text: '改过了\n两段', done: true });
  });

  it('原来没有批注的文档：新建 comments.xml（与扩展部件），登记内容类型与关系；点批注落在字缝里', async () => {
    const d = await UltimateWord.load(commentDocx(true));
    expect(d.comments()).toEqual([]);
    const hit = d.find('乙')[0];
    if (hit === undefined) throw new Error('找不到「乙」');
    const id = d.addComment({ start: hit.end, end: hit.end }, '插一句', {
      author: '孙八',
      date: '2026-10-12T12:00:00Z',
    });
    expect(id).toBe('0');
    d.resolveComment(id);
    const out = await bytesOf(d);
    const parts = unzip(out);
    expect(text(parts, '[Content_Types].xml')).toContain(
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>',
    );
    expect(text(parts, 'word/_rels/document.xml.rels')).toContain('Target="commentsExtended.xml"');
    const comments = text(parts, 'word/comments.xml');
    expect(comments).toMatch(
      /^<\?xml[^>]*\?><w:comments xmlns:w="[^"]+" xmlns:w14="[^"]+" xmlns:mc="[^"]+" mc:Ignorable="w14">/,
    );
    expect(text(parts, 'word/document.xml')).toContain(
      '<w:r><w:t>甲乙</w:t></w:r><w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/>' +
        '<w:r><w:commentReference w:id="0"/></w:r><w:r><w:t>丙</w:t></w:r>',
    );
    const again = await UltimateWord.load(out);
    expect(pick(again)).toEqual([['0', '', undefined, true, '插一句', '孙八']]);
    const range = again.comments()[0]?.range;
    expect(range?.start).toEqual(range?.end);
    expect(range?.start).toEqual(again.find('乙')[0]?.end);
  });
});

describe('批注的编辑：没有扩展部件的文件（WPS / 老版 Word）', () => {
  /** 同一份样本去掉 `w14:paraId` 与 commentsExtended.xml */
  function wpsDocx(): Uint8Array {
    const files = unzip(commentDocx());
    const enc = new TextEncoder();
    files.set(
      'word/comments.xml',
      enc.encode(text(files, 'word/comments.xml').replace(/ w14:paraId="[^"]*"/g, '')),
    );
    files.delete('word/commentsExtended.xml');
    files.set(
      'word/_rels/document.xml.rels',
      enc.encode(text(files, 'word/_rels/document.xml.rels').replace(/<Relationship Id="rId2"[^>]*>/, '')),
    );
    files.set(
      '[Content_Types].xml',
      enc.encode(
        text(files, '[Content_Types].xml').replace(/<Override PartName="\/word\/commentsExtended[^>]*>/, ''),
      ),
    );
    return zip(files);
  }

  it('只改正文时不新建扩展部件；一解决就给每条批注补 paraId（全文唯一）并新建扩展部件', async () => {
    const d = await UltimateWord.load(wpsDocx());
    expect(d.comments().every((c) => !c.done && c.parent === undefined)).toBe(true);
    d.editComment('2', '补两句');
    let parts = unzip(await bytesOf(d));
    expect(parts.has('word/commentsExtended.xml')).toBe(false);
    expect(text(parts, 'word/comments.xml')).not.toContain('w14:paraId');

    d.resolveComment('1');
    parts = unzip(await bytesOf(d));
    const comments = text(parts, 'word/comments.xml');
    const ids = [...comments.matchAll(/w14:paraId="([0-9A-F]{8})"/g)].map((m) => m[1]);
    expect(ids).toHaveLength(4);
    expect(new Set(ids).size).toBe(4);
    expect(text(parts, 'word/commentsExtended.xml').match(/<w15:commentEx /g)).toHaveLength(4);
    const again = await UltimateWord.load(await bytesOf(d));
    expect(again.comments().map((c) => [c.id, c.done, c.text])).toEqual([
      ['0', false, '这里要改'],
      ['3', false, '同意'],
      ['1', true, '第一段\n第二段'],
      ['2', false, '补两句'],
    ]);
  });
});
