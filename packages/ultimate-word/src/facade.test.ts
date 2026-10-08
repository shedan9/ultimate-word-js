/**
 * 门面的判据不是真值残差 —— 字摆在哪由底下已标定完的几层决定，这里只验「接线没接错」：
 * 四种 `LoadSource` 都通到同一份布局、`find` 自动带上域求值结果、`rangeOf` 认 id 也认节点、
 * 随库度量包在模块加载时就已经在注册表里。跑在纯 Node 上，`mount()` 的那一半在
 * `apps/playground/tests/facade.html`。
 */
import { readFileSync } from 'node:fs';
import { FALLBACK_METRICS, FontRegistry } from '@uw/fonts';
import { unzip, zip } from '@uw/ooxml';
import { beforeAll, describe, expect, it } from 'vitest';
import type { UwDocument } from './index.ts';
import { DOCX_MIME, UltimateWord, UwError, UwErrorCode } from './index.ts';

const FIXTURE = new URL('../../../apps/fidelity/fixtures/gongwen-01.docx', import.meta.url);
const TRUTH = new URL('../../../apps/fidelity/fixtures/gongwen-01.truth.json', import.meta.url);

const bytes = new Uint8Array(readFileSync(FIXTURE));
const truth = JSON.parse(readFileSync(TRUTH, 'utf8')) as { pageCount: number };

let doc: UwDocument;
beforeAll(async () => {
  doc = await UltimateWord.load(bytes);
});

describe('UltimateWord.load', () => {
  it('页数与 Word 真值一致 —— 说明六个包接成了 playground 那条链', () => {
    expect(doc.pageCount).toBe(truth.pageCount);
    expect(doc.layout.pages).toHaveLength(truth.pageCount);
  });

  it('ArrayBuffer / Blob / Response 三种来源与 Uint8Array 得到同一份布局', async () => {
    const fromBuffer = await UltimateWord.load(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    );
    const fromBlob = await UltimateWord.load(new Blob([bytes]));
    const fromResponse = await UltimateWord.load(new Response(bytes));
    for (const d of [fromBuffer, fromBlob, fromResponse]) {
      expect(d.pageCount).toBe(doc.pageCount);
      expect(d.layout).toEqual(doc.layout);
    }
  });

  it('不是 zip 时抛 UwError（结构性错误），不是记诊断', async () => {
    await expect(UltimateWord.load(new Uint8Array([1, 2, 3]))).rejects.toSatisfy(
      (e: unknown) => e instanceof UwError && e.code === UwErrorCode.NOT_A_ZIP,
    );
  });

  it('响应不成功的 Response 直接拒绝，不拿 404 页面去解 zip', async () => {
    await expect(UltimateWord.load(new Response('nope', { status: 404 }))).rejects.toThrow(/404/);
  });

  it('已中止的 signal 拒绝加载', async () => {
    const ctl = new AbortController();
    ctl.abort();
    await expect(UltimateWord.load(new Blob([bytes]), { signal: ctl.signal })).rejects.toThrow();
  });

  it('诊断是纯数据数组（可能为空）', () => {
    expect(Array.isArray(doc.diagnostics)).toBe(true);
    for (const d of doc.diagnostics) expect(typeof d.code).toBe('string');
  });
});

describe('UwDocument 的查询', () => {
  it('find 跨 run 找到真值里的那一行标题', () => {
    const hits = doc.find('保真度验证');
    expect(hits.length).toBeGreaterThan(0);
    for (const r of hits) expect(doc.compare(r.start, r.end)).toBe(-1);
  });

  it('find 的 limit / matchCase 透传', () => {
    expect(doc.find('通知', { limit: 1 })).toHaveLength(1);
    expect(doc.find(/word/, { matchCase: true }).length).toBeLessThanOrEqual(doc.find('word').length);
  });

  it('query 走可编辑的那棵树，rangeOf 认 id 也认节点', () => {
    const paragraphs = doc.query('paragraph');
    expect(paragraphs.length).toBeGreaterThan(0);
    const withRuns = paragraphs.find((p) => p.kind === 'paragraph' && p.runs.length > 0);
    if (withRuns === undefined) throw new Error('fixture 里没有带 run 的段落');
    const byNode = doc.rangeOf(withRuns);
    const byId = doc.rangeOf(withRuns.id);
    expect(byNode).toBeDefined();
    expect(byId).toEqual(byNode);
    expect(doc.rangeOf('不存在的 id')).toBeUndefined();
  });

  it('query 对 image / field / sdt 抛错而不是答空', () => {
    expect(() => doc.query('image')).toThrow();
  });

  it('compare 对不属于这份文档的位置抛错', () => {
    const r = doc.find('通知')[0];
    if (r === undefined) throw new Error('fixture 里没有「通知」');
    expect(doc.compare(r.start, r.start)).toBe(0);
    expect(() => doc.compare(r.start, { nodeId: 'r-not-here', contentIndex: 0, offset: 0 })).toThrow();
  });
});

describe('UltimateWord.fonts', () => {
  it('随库 17 款度量包在模块加载时就已注册', () => {
    expect(UltimateWord.fonts.status('宋体')).toBe('metrics');
    expect(UltimateWord.fonts.status('Times New Roman')).toBe('metrics');
    expect(UltimateWord.fonts.registry.families()).toHaveLength(17);
  });

  it('替换表把 missing 变成 fallback', () => {
    expect(UltimateWord.fonts.status('仿宋_GB2312')).toBe('missing');
    UltimateWord.fonts.substitute({ 仿宋_GB2312: '仿宋' });
    expect(UltimateWord.fonts.status('仿宋_GB2312')).toBe('fallback');
  });

  it('register 走动态 import 的 decode，字节不对时拒绝而不是静默注册', async () => {
    await expect(UltimateWord.fonts.register('坏字体', new Uint8Array([0, 1, 2, 3]))).rejects.toThrow();
    expect(UltimateWord.fonts.status('坏字体')).toBe('missing');
  });
});

describe('门面编辑复用段落缓存', () => {
  it('事务后可查询新增文字，撤销恢复原布局，重做恢复编辑布局', async () => {
    const editing = await UltimateWord.load(bytes);
    const initial = structuredClone(editing.layout);
    const pos = editing.find('通知')[0]?.start;
    if (!pos) throw new Error('样本缺少通知');
    editing.tx((tx) => {
      tx.insertText(pos, '缓存回归内容'.repeat(40));
    });
    const changed = structuredClone(editing.layout);
    expect(editing.find('缓存回归内容')).toHaveLength(40);
    expect(changed).not.toEqual(initial);
    editing.undo();
    expect(editing.layout).toEqual(initial);
    expect(editing.find('缓存回归内容')).toHaveLength(0);
    editing.redo();
    expect(editing.layout).toEqual(changed);
  });

  it('字体注册后下一次事务与重新加载使用同一度量，不复用旧布局', async () => {
    const fonts = new FontRegistry();
    const editing = await UltimateWord.load(bytes, { fonts });
    const initial = structuredClone(editing.layout);
    const pos = editing.find('通知')[0]?.start;
    if (!pos) throw new Error('样本缺少通知');
    fonts.register('仿宋', { kind: 'file', metrics: FALLBACK_METRICS, advance: () => 1600 });
    editing.tx((tx) => {
      tx.insertText(pos, '刷新');
    });
    editing.undo();
    const fresh = await UltimateWord.load(bytes, { fonts });
    expect(editing.layout).toEqual(fresh.layout);
    expect(editing.layout).not.toEqual(initial);
  });
});

describe('doc.toDocx', () => {
  it('没编辑：导出的就是原文件的每个部件，MIME 是 docx', async () => {
    const fresh = await UltimateWord.load(bytes);
    const blob = await fresh.toDocx();
    expect(blob.type).toBe(DOCX_MIME);
    const before = unzip(bytes);
    const after = unzip(new Uint8Array(await blob.arrayBuffer()));
    for (const [k, v] of before) expect(after.get(k), k).toEqual(v);
  });

  it('编辑后导出再加载：改动在，页数与编辑后的布局一致', async () => {
    const editing = await UltimateWord.load(bytes);
    const pos = editing.find('通知')[0]?.start;
    if (!pos) throw new Error('样本缺少通知');
    editing.tx((tx) => {
      tx.insertText(pos, '导出回归');
    });
    const reloaded = await UltimateWord.load(await editing.toDocx());
    expect(reloaded.find('导出回归通知')).toHaveLength(1);
    expect(reloaded.layout).toEqual(editing.layout);
  });
});

describe('UwDocument 的事件', () => {
  it('事务 / 撤销 / 重做各派发一次 document:change，之后紧跟 layout:done', async () => {
    const editing = await UltimateWord.load(bytes);
    const log: string[] = [];
    let seen: { pageCount: number; duration: number; iterations: number } | undefined;
    const a = editing.on('document:change', ({ changeSet, source }) => {
      // 监听者里读到的已经是新布局
      log.push(`${source}:${changeSet.changes.length > 0}:${editing.find('事件回归').length}`);
    });
    const b = editing.on('layout:done', (info) => {
      log.push('layout');
      seen = info;
    });
    const pos = editing.find('通知')[0]?.start;
    if (!pos) throw new Error('样本缺少通知');
    editing.tx((tx) => {
      tx.insertText(pos, '事件回归');
    });
    editing.undo();
    editing.redo();
    expect(log).toEqual(['tx:true:1', 'layout', 'undo:true:0', 'layout', 'redo:true:1', 'layout']);
    expect(seen?.pageCount).toBe(editing.pageCount);
    expect(seen?.iterations).toBeGreaterThanOrEqual(1);
    expect(seen?.duration).toBeGreaterThanOrEqual(0);
    // 没有修改的事务不派发；dispose 之后不再收到
    editing.tx(() => undefined);
    a.dispose();
    b.dispose();
    editing.undo();
    expect(log).toHaveLength(6);
  });

  it('一个监听者抛错不挡后面的监听者，也不让事务半途而废', async () => {
    const editing = await UltimateWord.load(bytes);
    const original = globalThis.reportError;
    const reported: unknown[] = [];
    globalThis.reportError = (e: unknown) => reported.push(e);
    try {
      let after = 0;
      editing.on('document:change', () => {
        throw new Error('宿主的 bug');
      });
      editing.on('document:change', () => {
        after++;
      });
      const pos = editing.find('通知')[0]?.start;
      if (!pos) throw new Error('样本缺少通知');
      expect(() =>
        editing.tx((tx) => {
          tx.insertText(pos, '抛错');
        }),
      ).not.toThrow();
      expect(after).toBe(1);
      expect(reported).toHaveLength(1);
      expect(editing.find('抛错通知')).toHaveLength(1);
    } finally {
      globalThis.reportError = original;
    }
  });

  it('重排新发现的缺字体报一次 diagnostic 并追加进 diagnostics，同一条不重复报', async () => {
    const editing = await UltimateWord.load(bytes);
    const before = editing.diagnostics.length;
    const got: string[] = [];
    editing.on('diagnostic', (d) => got.push(`${d.code}:${d.message}`));
    const hit = editing.find('通知')[0];
    if (!hit) throw new Error('样本缺少通知');
    const missing = { ascii: '事件回归缺字体', hAnsi: '事件回归缺字体', eastAsia: '事件回归缺字体' };
    editing.tx((tx) => {
      tx.setRunProps(hit, { fonts: missing });
    });
    expect(got).toHaveLength(1);
    expect(got[0]).toMatch(/^font-missing:.*事件回归缺字体/);
    expect(editing.diagnostics).toHaveLength(before + 1);
    // 撤销再重做走的是同一款缺失字体：度量器不再报，门面也不会重复
    editing.undo();
    editing.redo();
    expect(got).toHaveLength(1);
  });
});

describe('doc.bindings', () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const control = (tag: string, text: string, pr = '') =>
    `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${pr}</w:sdtPr><w:sdtContent><w:r><w:t>${text}</w:t></w:r></w:sdtContent></w:sdt>`;
  /** 最小模板：两个同名坑位、一个占位状态的纯文本、一个下拉框、一个复选框 */
  function template(): Uint8Array {
    const enc = new TextEncoder();
    const body =
      `<w:p><w:r><w:t>申请人：</w:t></w:r>${control('applicant', '单击输入', '<w:showingPlcHdr/><w:text/>')}</w:p>` +
      `<w:p><w:r><w:t>日期：</w:t></w:r>${control('date', '某日')}<w:r><w:t>，签名：</w:t></w:r>${control('applicant', '单击输入', '<w:showingPlcHdr/><w:text/>')}</w:p>` +
      `<w:p>${control('level', '一般', '<w:dropDownList><w:listItem w:displayText="紧急" w:value="1"/><w:listItem w:displayText="一般" w:value="2"/></w:dropDownList>')}${control('agree', '☐', '<w14:checkbox/>')}</w:p>`;
    return zip(
      new Map([
        [
          '[Content_Types].xml',
          enc.encode(
            '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
              '<Default Extension="xml" ContentType="application/xml"/>' +
              '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
          ),
        ],
        [
          '_rels/.rels',
          enc.encode(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
          ),
        ],
        [
          'word/document.xml',
          enc.encode(
            `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`,
          ),
        ],
      ]),
    );
  }

  it('list 按文档序列出坑位与现值；set 只记下，apply 一次提交成一个撤销单元', async () => {
    const d = await UltimateWord.load(template());
    expect(d.bindings.list().map((b) => [b.key, b.value, b.placeholder])).toEqual([
      ['applicant', '单击输入', true],
      ['date', '某日', false],
      ['applicant', '单击输入', true],
      ['level', '一般', false],
      ['agree', '☐', false],
    ]);
    expect(d.bindings.list()[3]?.options).toEqual(['紧急', '一般']);
    const changes: string[] = [];
    d.on('document:change', (e) => changes.push(e.source));
    d.bindings.set('applicant', '张三').set('level', '1');
    expect(d.find('张三')).toHaveLength(0);
    const { skipped } = d.bindings.setMany({ date: '2026年9月28日', unused: 'x' });
    expect(skipped).toEqual(['unused']);
    expect(d.bindings.apply()).toBeDefined();
    expect(changes).toEqual(['tx']);
    expect(d.bindings.pending.size).toBe(0);
    expect(d.find('张三')).toHaveLength(2);
    expect(d.bindings.list().map((b) => b.value)).toEqual(['张三', '2026年9月28日', '张三', '紧急', '☐']);
    expect(d.bindings.list().every((b) => !b.placeholder)).toBe(true);
    d.undo();
    expect(d.bindings.list().map((b) => b.value)).toEqual(['单击输入', '某日', '单击输入', '一般', '☐']);
  });

  it('名字拼错、复选框在 set 时就抛；下拉框没有的选项在 apply 时整批回滚、值留着', async () => {
    const d = await UltimateWord.load(template());
    expect(() => d.bindings.set('aplicant', '张三')).toThrow(/没有坑位/);
    expect(() => d.bindings.set('agree', '☑')).toThrow(/checkbox/);
    d.bindings.set('applicant', '张三').set('level', '特急');
    expect(() => d.bindings.apply()).toThrow();
    expect(d.find('张三')).toHaveLength(0);
    expect(d.canUndo).toBe(false);
    expect(d.bindings.pending.get('applicant')).toBe('张三');
    d.bindings.discard();
    expect(d.bindings.apply()).toBeUndefined();
  });

  it('填完导出再加载：坑位还是坑位，值留着、占位标记没了', async () => {
    const d = await UltimateWord.load(template());
    d.bindings.set('applicant', '李四');
    d.bindings.apply();
    const again = await UltimateWord.load(await d.toDocx());
    const applicant = again.bindings.list().filter((b) => b.key === 'applicant');
    expect(applicant.map((b) => [b.value, b.placeholder])).toEqual([
      ['李四', false],
      ['李四', false],
    ]);
  });
});

describe('doc.replaceAll', () => {
  it('一个事务换掉全部命中，一次撤销全回来；字符串查找时 $ 不展开', async () => {
    const editing = await UltimateWord.load(bytes);
    const initial = structuredClone(editing.layout);
    const count = editing.find('通知').length;
    expect(count).toBeGreaterThan(0);
    let changes = 0;
    editing.on('document:change', () => changes++);
    expect(editing.replaceAll('通知', '$&告')).toEqual({ replaced: count, skipped: 0 });
    expect(changes).toBe(1);
    expect(editing.find('$&告')).toHaveLength(count);
    expect(editing.find('通知')).toHaveLength(0);
    const reloaded = await UltimateWord.load(await editing.toDocx());
    expect(reloaded.find('$&告')).toHaveLength(count);
    editing.undo();
    expect(editing.layout).toEqual(initial);
  });

  it('正则替换展开捕获组，函数替换拿到命中；没有命中不开事务', async () => {
    const editing = await UltimateWord.load(bytes);
    const count = editing.find('通知').length;
    expect(editing.replaceAll(/通(知)/, '[$1]')).toEqual({ replaced: count, skipped: 0 });
    expect(editing.find('[知]')).toHaveLength(count);
    expect(editing.replaceAll('[知]', (m) => `${m.text.length}`).replaced).toBe(count);
    expect(editing.canUndo).toBe(true);
    editing.undo();
    expect(editing.find('[知]')).toHaveLength(count);
    expect(editing.replaceAll('不存在的文字', 'x')).toEqual({ replaced: 0, skipped: 0 });
    // 回调抛错时一个字都没改
    expect(() =>
      editing.replaceAll('[知]', () => {
        throw new Error('停');
      }),
    ).toThrow('停');
    expect(editing.find('[知]')).toHaveLength(count);
  });
});

describe('目录页码（PAGEREF）与书签跳转', () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const fld = (type: 'begin' | 'separate' | 'end') => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
  const instr = (text: string) => `<w:r><w:instrText xml:space="preserve"> ${text} </w:instrText></w:r>`;
  /** Word 生成的目录条目长这样：HYPERLINK 包着标题文字 + 制表位 + 嵌套的 PAGEREF。存着的页码故意写错成 9 */
  const entry = (name: string, title: string) =>
    `${fld('begin')}${instr(`HYPERLINK \\l "${name}"`)}${fld('separate')}<w:r><w:t>${title}</w:t></w:r><w:r><w:tab/></w:r>` +
    `${fld('begin')}${instr(`PAGEREF ${name} \\h`)}${fld('separate')}<w:r><w:t>9</w:t></w:r>${fld('end')}${fld('end')}`;
  const heading = (name: string, title: string) =>
    `<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:bookmarkStart w:id="0" w:name="${name}"/><w:r><w:t>${title}</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>`;

  /** 第 1 页目录两条，第一章在第 2 页、第二章在第 3 页 */
  function tocDocx(): Uint8Array {
    const enc = new TextEncoder();
    const body =
      `<w:p>${fld('begin')}${instr('TOC \\o "1-3" \\h \\z \\u')}${fld('separate')}${entry('_Toc1', '第一章')}</w:p>` +
      `<w:p>${entry('_Toc2', '第二章')}${fld('end')}</w:p>` +
      `<w:p><w:r><w:t>前言</w:t></w:r></w:p>${heading('_Toc1', '第一章')}${heading('_Toc2', '第二章')}`;
    return zip(
      new Map([
        [
          '[Content_Types].xml',
          enc.encode(
            '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
              '<Default Extension="xml" ContentType="application/xml"/>' +
              '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
          ),
        ],
        [
          '_rels/.rels',
          enc.encode(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
          ),
        ],
        [
          'word/document.xml',
          enc.encode(
            `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`,
          ),
        ],
      ]),
    );
  }

  /** 第一页每一行的文字 —— 目录条目「标题 + 页码」 */
  function tocLines(d: UwDocument): string[] {
    const out: string[] = [];
    for (const block of d.layout.pages[0]?.blocks ?? []) {
      if (block.kind !== 'paragraph') continue;
      for (const placed of block.lines) out.push(placed.line.fragments.map((f) => f.text).join(''));
    }
    return out;
  }

  it('打开时就把存着的错页码算对，编辑把标题挤到后一页时跟着变，撤销变回来', async () => {
    const d = await UltimateWord.load(tocDocx());
    expect(d.pageCount).toBe(3);
    expect(tocLines(d).slice(0, 2)).toEqual(['第一章2', '第二章3']);

    const preface = d.find('前言')[0];
    if (preface === undefined) throw new Error('找不到「前言」');
    d.tx((t) => {
      t.insertInline(preface.start, 'pageBreak');
    });
    expect(d.pageCount).toBe(4);
    expect(tocLines(d).slice(0, 2)).toEqual(['第一章3', '第二章4']);
    d.undo();
    expect(tocLines(d).slice(0, 2)).toEqual(['第一章2', '第二章3']);
  });

  it('rangeOfBookmark 给出标题段落的 range（目录跳转的落点），不存在答 undefined', async () => {
    const d = await UltimateWord.load(tocDocx());
    const range = d.rangeOfBookmark('_Toc2');
    const title = d.find('第二章').at(-1);
    expect(range).toBeDefined();
    expect(title).toBeDefined();
    if (range === undefined || title === undefined) return;
    expect(d.compare(range.start, title.start)).toBe(0);
    expect(d.rangeOfBookmark('_Toc404')).toBeUndefined();
  });
});

describe('题注编号（SEQ）与章节引用（STYLEREF）', () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const fld = (type: 'begin' | 'separate' | 'end') => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
  /** Word「插入题注 → 包含章节号」写出来的是 `SEQ 图 \* ARABIC \s 1`。存着的号故意全写成 9 */
  const caption = `<w:p><w:r><w:t xml:space="preserve">图 </w:t></w:r>${fld('begin')}<w:r><w:instrText xml:space="preserve"> SEQ 图 \\* ARABIC \\s 1 </w:instrText></w:r>${fld('separate')}<w:r><w:t>9</w:t></w:r>${fld('end')}</w:p>`;
  const para = (text: string, style?: string) =>
    `<w:p>${style === undefined ? '' : `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>`}<w:r><w:t>${text}</w:t></w:r></w:p>`;

  /** 标题样式的 id 是 `1`（中文版 Word 就这么写），大纲级别只写在样式里 —— 段落自己一个字都不提 */
  function seqDocx(
    body = `${para('第一章', '1')}${caption}${caption}${para('第二章')}${caption}`,
  ): Uint8Array {
    const enc = new TextEncoder();
    return zip(
      new Map([
        [
          '[Content_Types].xml',
          enc.encode(
            '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
              '<Default Extension="xml" ContentType="application/xml"/>' +
              '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
              '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
          ),
        ],
        [
          '_rels/.rels',
          enc.encode(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
          ),
        ],
        [
          'word/_rels/document.xml.rels',
          enc.encode(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/></Relationships>`,
          ),
        ],
        [
          'word/styles.xml',
          enc.encode(
            `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>`,
          ),
        ],
        [
          'word/document.xml',
          enc.encode(
            `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`,
          ),
        ],
      ]),
    );
  }

  /** 第一页所有以「图」开头的行 */
  function captions(d: UwDocument): string[] {
    const out: string[] = [];
    for (const block of d.layout.pages[0]?.blocks ?? []) {
      if (block.kind !== 'paragraph') continue;
      for (const placed of block.lines) {
        const text = placed.line.fragments.map((f) => f.text).join('');
        if (text.startsWith('图')) out.push(text);
      }
    }
    return out;
  }

  it('打开时就按文档序重数；把「第二章」设成标题 1 后从 1 数起，撤销变回来', async () => {
    const d = await UltimateWord.load(seqDocx());
    expect(captions(d)).toEqual(['图 1', '图 2', '图 3']);

    const chapter = d.find('第二章')[0];
    if (chapter === undefined) throw new Error('找不到「第二章」');
    d.tx((t) => {
      t.setParagraphProps(chapter, { styleId: '1' });
    });
    expect(captions(d)).toEqual(['图 1', '图 2', '图 1']);
    d.undo();
    expect(captions(d)).toEqual(['图 1', '图 2', '图 3']);
  });

  it('STYLEREF 按中文界面名认出英文 w:name 的内建标题；套上标题样式后跟着变，撤销变回来', async () => {
    const ref = `<w:p><w:r><w:t>见</w:t></w:r>${fld('begin')}<w:r><w:instrText xml:space="preserve"> STYLEREF "标题 1" </w:instrText></w:r>${fld('separate')}<w:r><w:t>旧</w:t></w:r>${fld('end')}</w:p>`;
    const d = await UltimateWord.load(seqDocx(`${para('第一章', '1')}${para('第二章')}${ref}`));
    const line = () => d.layout.pages[0]?.blocks.at(-1);
    const text = (): string => {
      const block = line();
      if (block === undefined || block.kind !== 'paragraph') return '';
      return block.lines.map((l) => l.line.fragments.map((f) => f.text).join('')).join('');
    };
    expect(text()).toBe('见第一章');

    const chapter = d.find('第二章')[0];
    if (chapter === undefined) throw new Error('找不到「第二章」');
    d.tx((t) => {
      t.setParagraphProps(chapter, { styleId: '1' });
    });
    expect(text()).toBe('见第二章');
    d.undo();
    expect(text()).toBe('见第一章');
  });

  it('DATE 打开即显示当天（按 \\@ 格式串），没写格式串的照旧', async () => {
    const date = (picture: string) =>
      `<w:p>${fld('begin')}<w:r><w:instrText xml:space="preserve"> DATE ${picture} </w:instrText></w:r>${fld('separate')}<w:r><w:t>1999</w:t></w:r>${fld('end')}</w:p>`;
    const before = new Date().getFullYear();
    const d = await UltimateWord.load(seqDocx(`${date('\\@ "yyyy"')}${date('')}`));
    const after = new Date().getFullYear();
    const texts = (d.layout.pages[0]?.blocks ?? []).map((b) =>
      b.kind === 'paragraph' ? b.lines.map((l) => l.line.fragments.map((f) => f.text).join('')).join('') : '',
    );
    expect([String(before), String(after)]).toContain(texts[0]);
    expect(texts[1]).toBe('1999');
  });
});

describe('更新目录（TOC 重新生成）', () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const fld = (type: 'begin' | 'separate' | 'end') => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
  const para = (text: string, style?: string) =>
    `<w:p>${style === undefined ? '' : `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>`}<w:r><w:t>${text}</w:t></w:r></w:p>`;
  const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

  /** 一份陈旧的目录（只有一条早就删掉的「旧条目」），后面第 2 页第一章、第 3 页第二章，标题都没有书签 */
  function tocDocx(): Uint8Array {
    const enc = new TextEncoder();
    const body =
      `<w:p>${fld('begin')}<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>${fld('separate')}` +
      `<w:r><w:t>旧条目</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>9</w:t></w:r></w:p><w:p>${fld('end')}</w:p>` +
      `${para('前言')}${pageBreak}${para('第一章', '1')}${pageBreak}${para('第二章', '1')}`;
    return zip(
      new Map([
        [
          '[Content_Types].xml',
          enc.encode(
            '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
              '<Default Extension="xml" ContentType="application/xml"/>' +
              '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
              '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
          ),
        ],
        [
          '_rels/.rels',
          enc.encode(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
          ),
        ],
        [
          'word/_rels/document.xml.rels',
          enc.encode(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/></Relationships>`,
          ),
        ],
        [
          'word/styles.xml',
          enc.encode(
            `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>` +
              '<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>',
          ),
        ],
        [
          'word/document.xml',
          enc.encode(
            `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`,
          ),
        ],
      ]),
    );
  }

  /** 第一页上目录那几行（到「前言」为止） */
  function tocLines(d: UwDocument): string[] {
    const out: string[] = [];
    for (const block of d.layout.pages[0]?.blocks ?? []) {
      if (block.kind !== 'paragraph') continue;
      for (const placed of block.lines) out.push(placed.line.fragments.map((f) => f.text).join(''));
    }
    return out.slice(0, out.indexOf('前言'));
  }

  it('按现在的标题重新生成条目、补书签，页码按标题所在页；加标题再更新多一条，撤销退回', async () => {
    const d = await UltimateWord.load(tocDocx());
    expect(tocLines(d)).toEqual(['旧条目9', '']);

    expect(d.updateTableOfContents()).toEqual({ updated: 1, skipped: [] });
    expect(tocLines(d)).toEqual(['第一章2', '第二章3', '']);
    // 条目是跳到标题的超链接，书签是新补的
    const target = d.rangeOfBookmark('_Toc100000001');
    const second = d.find('第二章').at(-1);
    if (target === undefined || second === undefined) throw new Error('第二章的书签没补上');
    expect(d.compare(target.start, second.start)).toBe(0);

    // 把「前言」也设成标题：目录不自己变（Word 也不），更新之后才多一条
    const preface = d.find('前言')[0];
    if (preface === undefined) throw new Error('找不到「前言」');
    d.tx((t) => {
      t.setParagraphProps(preface, { styleId: '1' });
    });
    expect(tocLines(d)).toEqual(['第一章2', '第二章3', '']);
    d.updateTableOfContents();
    expect(tocLines(d)).toEqual(['前言1', '第一章2', '第二章3', '']);
    d.undo();
    expect(tocLines(d)).toEqual(['第一章2', '第二章3', '']);
  });

  it('导出的 docx 里存着新条目与页码，重新打开一样；图表目录跳过不动', async () => {
    const d = await UltimateWord.load(tocDocx());
    d.updateTableOfContents();
    const again = await UltimateWord.load(await d.toDocx());
    expect(tocLines(again)).toEqual(['第一章2', '第二章3', '']);
    expect(again.rangeOfBookmark('_Toc100000000')).toBeDefined();
    const xml = new TextDecoder().decode(
      unzip(new Uint8Array(await (await d.toDocx()).arrayBuffer())).get('word/document.xml'),
    );
    expect(xml).toContain('<w:tab w:val="right" w:leader="dot" w:pos="8306"/>');
    expect(xml).toMatch(
      /PAGEREF _Toc100000001 \\h <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:r><w:t>3<\/w:t>/,
    );

    const figures = await UltimateWord.load(
      new Uint8Array(
        zip(
          new Map(
            [...unzip(tocDocx())].map(([k, v]) => [
              k,
              k === 'word/document.xml'
                ? new TextEncoder().encode(new TextDecoder().decode(v).replace('\\o "1-3"', '\\c "图"'))
                : v,
            ]),
          ),
        ),
      ),
    );
    expect(figures.updateTableOfContents()).toEqual({
      updated: 0,
      skipped: [{ reason: expect.stringContaining('图表目录') }],
    });
    expect(figures.canUndo).toBe(false);
  });
});

describe('脚注', () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ref = (id: string) =>
    `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
  const note = (id: string, text: string, more = 0) =>
    `<w:footnote w:id="${id}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>` +
    `<w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p>` +
    Array.from({ length: more }, (_, i) => `<w:p><w:r><w:t>续${i + 1}</w:t></w:r></w:p>`).join('') +
    '</w:footnote>';

  /**
   * 两段各引一条；`w:id` 与出现顺序相反（2 在前），号必须按先后数。格式在 settings.xml 里设成圈码。
   * `more` 给甲注（id 2）再续这么多段，用来造长过一页的脚注
   */
  function footnoteDocx(more = 0, restart = ''): Uint8Array {
    const enc = new TextEncoder();
    const body = `<w:p><w:r><w:t>甲</w:t></w:r>${ref('2')}</w:p><w:p><w:r><w:t>乙</w:t></w:r>${ref('1')}</w:p>`;
    const ct = (part: string, type: string) =>
      `<Override PartName="/word/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`;
    const rel = (id: string, type: string, target: string) =>
      `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
    return zip(
      new Map([
        [
          '[Content_Types].xml',
          enc.encode(
            '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
              '<Default Extension="xml" ContentType="application/xml"/>' +
              ct('document.xml', 'document.main') +
              ct('footnotes.xml', 'footnotes') +
              ct('settings.xml', 'settings') +
              '</Types>',
          ),
        ],
        [
          '_rels/.rels',
          enc.encode(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId1', 'officeDocument', 'word/document.xml')}</Relationships>`,
          ),
        ],
        [
          'word/_rels/document.xml.rels',
          enc.encode(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId1', 'footnotes', 'footnotes.xml')}${rel('rId2', 'settings', 'settings.xml')}</Relationships>`,
          ),
        ],
        [
          'word/settings.xml',
          enc.encode(
            `<w:settings xmlns:w="${W}"><w:footnotePr><w:numFmt w:val="decimalEnclosedCircleChinese"/>${restart === '' ? '' : `<w:numRestart w:val="${restart}"/>`}<w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr></w:settings>`,
          ),
        ],
        [
          'word/footnotes.xml',
          enc.encode(
            `<w:footnotes xmlns:w="${W}">` +
              '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>' +
              '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
              `${note('1', '乙注')}${note('2', '甲注', more)}</w:footnotes>`,
          ),
        ],
        [
          'word/document.xml',
          enc.encode(
            `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
          ),
        ],
      ]),
    );
  }

  /** 正文里每一段的号（域结果片段）与脚注区每一条的文字 */
  function marks(d: UwDocument): { body: string[]; notes: string[] } {
    const page = d.layout.pages[0];
    const body: string[] = [];
    for (const block of page?.blocks ?? []) {
      if (block.kind !== 'paragraph') continue;
      for (const placed of block.lines) for (const f of placed.line.fragments) if (f.field) body.push(f.text);
    }
    const notes: string[] = [];
    for (const block of page?.footnotes?.blocks ?? []) {
      if (block.kind !== 'paragraph') continue;
      for (const placed of block.lines) notes.push(placed.line.fragments.map((f) => f.text).join(''));
    }
    return { body, notes };
  }

  it('号按引用先后数（不看 w:id）、用 settings 里的格式；页底按同一顺序排出内容，分隔线不报不认识的元素', async () => {
    const d = await UltimateWord.load(footnoteDocx());
    expect(marks(d)).toEqual({ body: ['①', '②'], notes: ['① 甲注', '② 乙注'] });
    expect(d.layout.pages[0]?.footnotes?.notes).toEqual(['footnote:2', 'footnote:1']);
    expect(d.diagnostics.filter((x) => x.code === 'unknown-element')).toEqual([]);
  });

  it('编辑后重排脚注仍在；导出重开布局一致，脚注部件原样', async () => {
    const d = await UltimateWord.load(footnoteDocx());
    const pos = d.find('乙')[0]?.start;
    if (!pos) throw new Error('样本缺少乙');
    d.tx((tx) => {
      tx.insertText(pos, '丙');
    });
    expect(marks(d)).toEqual({ body: ['①', '②'], notes: ['① 甲注', '② 乙注'] });
    const out = await d.toDocx();
    const reloaded = await UltimateWord.load(out);
    expect(reloaded.layout).toEqual(d.layout);
    const parts = unzip(new Uint8Array(await out.arrayBuffer()));
    expect(parts.get('word/footnotes.xml')).toEqual(unzip(footnoteDocx()).get('word/footnotes.xml'));
  });

  it('长过一页的脚注切开续排：续页先排剩下的那一截，后引的那条连同引用行跟到续页', async () => {
    const d = await UltimateWord.load(footnoteDocx(70));
    const pages = d.layout.pages;
    expect(pages).toHaveLength(2);
    const [a, b] = [pages[0]?.footnotes, pages[1]?.footnotes];
    expect(a?.notes).toEqual(['footnote:2']);
    expect(a?.continued).toBeUndefined();
    expect(b?.continued).toBe(true);
    // 续页：甲注剩下的那一截在前，乙注在后；引乙的那一行也在续页上
    expect(b?.notes.at(-1)).toBe('footnote:1');
    const texts = pages.flatMap((p) =>
      (p.footnotes?.blocks ?? []).flatMap((x) =>
        x.kind === 'paragraph'
          ? [x.lines.map((l) => l.line.fragments.map((f) => f.text).join('')).join('')]
          : [],
      ),
    );
    // 甲注 71 段（一页放不下）+ 乙注 1 段，每段恰好出现一次、先后不乱
    expect(texts).toEqual(['① 甲注', ...Array.from({ length: 70 }, (_, i) => `续${i + 1}`), '② 乙注']);
    const bodyText = (i: number) =>
      (pages[i]?.blocks ?? []).flatMap((x) =>
        x.kind === 'paragraph' ? x.lines.flatMap((l) => l.line.fragments.map((f) => f.text)) : [],
      );
    expect(bodyText(0)).toEqual(['甲', '①']);
    expect(bodyText(1)).toEqual(['乙', '②']);
    // 每一页的脚注区都在版心里，没有溢出
    for (const p of pages) {
      const c = p.geometry.content;
      if (p.footnotes !== undefined) expect(p.footnotes.y).toBeGreaterThanOrEqual(c.y);
    }
    expect(d.diagnostics.filter((x) => x.code === 'footnote-overflow')).toEqual([]);
  });

  it('settings 里设了每页重新编号：续页上后引的那条从 ① 起，正文与页底一致', async () => {
    const d = await UltimateWord.load(footnoteDocx(70, 'eachPage'));
    const pages = d.layout.pages;
    expect(pages).toHaveLength(2);
    const bodyMarks = (i: number) =>
      (pages[i]?.blocks ?? []).flatMap((x) =>
        x.kind === 'paragraph'
          ? x.lines.flatMap((l) => l.line.fragments.filter((f) => f.field).map((f) => f.text))
          : [],
      );
    expect(bodyMarks(0)).toEqual(['①']);
    expect(bodyMarks(1)).toEqual(['①']);
    const lastNote = pages[1]?.footnotes?.blocks.at(-1);
    const text =
      lastNote?.kind === 'paragraph' ? lastNote.lines[0]?.line.fragments.map((f) => f.text).join('') : '';
    expect(text).toBe('① 乙注');
    expect(d.diagnostics.map((x) => x.code)).not.toContain('note-restart-not-converged');
  });
});

describe('尾注', () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ref = (id: string) =>
    `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;
  const note = (id: string, text: string) =>
    `<w:endnote w:id="${id}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r>` +
    `<w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:endnote>`;

  /** 两段各引一条（`w:id` 与出现顺序相反），外加一条没人引的。格式没设 → 尾注默认小写罗马数字 */
  function endnoteDocx(): Uint8Array {
    const enc = new TextEncoder();
    const body = `<w:p><w:r><w:t>甲</w:t></w:r>${ref('2')}</w:p><w:p><w:r><w:t>乙</w:t></w:r>${ref('1')}</w:p>`;
    const ct = (part: string, type: string) =>
      `<Override PartName="/word/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`;
    const rel = (id: string, type: string, target: string) =>
      `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
    return zip(
      new Map([
        [
          '[Content_Types].xml',
          enc.encode(
            '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
              '<Default Extension="xml" ContentType="application/xml"/>' +
              ct('document.xml', 'document.main') +
              ct('endnotes.xml', 'endnotes') +
              '</Types>',
          ),
        ],
        [
          '_rels/.rels',
          enc.encode(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId1', 'officeDocument', 'word/document.xml')}</Relationships>`,
          ),
        ],
        [
          'word/_rels/document.xml.rels',
          enc.encode(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId1', 'endnotes', 'endnotes.xml')}</Relationships>`,
          ),
        ],
        [
          'word/endnotes.xml',
          enc.encode(
            `<w:endnotes xmlns:w="${W}">` +
              '<w:endnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:endnote>' +
              '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>' +
              `${note('1', '乙尾')}${note('2', '甲尾')}${note('3', '没人引')}</w:endnotes>`,
          ),
        ],
        [
          'word/document.xml',
          enc.encode(
            `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
          ),
        ],
      ]),
    );
  }

  /** 页上每一块的文字，尾注块前面标一个「尾:」 */
  function texts(d: UwDocument): string[] {
    const out: string[] = [];
    for (const block of d.layout.pages[0]?.blocks ?? []) {
      if (block.kind !== 'paragraph') continue;
      const text = block.lines.map((l) => l.line.fragments.map((f) => f.text).join('')).join('');
      out.push(block.endnote === undefined ? text : `尾:${text}`);
    }
    return out;
  }

  it('正文后面画分隔线、按引用先后排出尾注，号是小写罗马数字；没人引的不排、不再报「没排」', async () => {
    const d = await UltimateWord.load(endnoteDocx());
    expect(texts(d)).toEqual(['甲i', '乙ii', '尾:i 甲尾', '尾:ii 乙尾']);
    expect(d.layout.pages[0]?.noteSeparators?.map((s) => s.kind)).toEqual(['separator']);
    expect(d.diagnostics.map((x) => x.code)).not.toContain('endnotes-not-rendered');
    expect(d.diagnostics.filter((x) => x.code === 'unknown-element')).toEqual([]);
  });

  it('编辑后重排尾注仍在；导出重开布局一致，尾注部件原样', async () => {
    const d = await UltimateWord.load(endnoteDocx());
    const pos = d.find('乙')[0]?.start;
    if (!pos) throw new Error('样本缺少乙');
    d.tx((tx) => {
      tx.insertText(pos, '丙');
    });
    expect(texts(d)).toEqual(['甲i', '丙乙ii', '尾:i 甲尾', '尾:ii 乙尾']);
    const out = await d.toDocx();
    const reloaded = await UltimateWord.load(out);
    expect(reloaded.layout).toEqual(d.layout);
    const parts = unzip(new Uint8Array(await out.arrayBuffer()));
    expect(parts.get('word/endnotes.xml')).toEqual(unzip(endnoteDocx()).get('word/endnotes.xml'));
  });
});

describe('修订的显示', () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  /** 正文一处删除 + 一处插入（两位作者），页眉里也有一处插入 + 删除 —— 换视图要连页眉一起换 */
  function revisedDocx(): Uint8Array {
    const enc = new TextEncoder();
    const ct = (part: string, type: string) =>
      `<Override PartName="/word/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`;
    const rel = (id: string, type: string, target: string) =>
      `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
    const body =
      '<w:p><w:r><w:t>原文</w:t></w:r>' +
      '<w:del w:id="1" w:author="张三" w:date="2026-10-01T08:00:00Z"><w:r><w:delText>删掉</w:delText></w:r></w:del>' +
      '<w:ins w:id="2" w:author="李四" w:date="2026-10-02T08:00:00Z"><w:r><w:t>新增</w:t></w:r></w:ins>' +
      '<w:r><w:t>结尾</w:t></w:r></w:p>';
    const header =
      '<w:p><w:ins w:id="3" w:author="李四"><w:r><w:t>新眉</w:t></w:r></w:ins>' +
      '<w:del w:id="4" w:author="李四"><w:r><w:delText>旧眉</w:delText></w:r></w:del></w:p>';
    return zip(
      new Map([
        [
          '[Content_Types].xml',
          enc.encode(
            '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
              '<Default Extension="xml" ContentType="application/xml"/>' +
              ct('document.xml', 'document.main') +
              ct('header1.xml', 'header') +
              '</Types>',
          ),
        ],
        [
          '_rels/.rels',
          enc.encode(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId1', 'officeDocument', 'word/document.xml')}</Relationships>`,
          ),
        ],
        [
          'word/_rels/document.xml.rels',
          enc.encode(
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel('rId7', 'header', 'header1.xml')}</Relationships>`,
          ),
        ],
        ['word/header1.xml', enc.encode(`<w:hdr xmlns:w="${W}" xmlns:r="${R}">${header}</w:hdr>`)],
        [
          'word/document.xml',
          enc.encode(
            `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
          ),
        ],
      ]),
    );
  }

  /** 第一页正文与页眉画出来的字（各自拼成一串）+ 正文片段的颜色 */
  function painted(d: UwDocument): { body: string; header: string; colors: string[] } {
    const page = d.layout.pages[0];
    const texts = (blocks: NonNullable<typeof page>['blocks']) =>
      blocks.flatMap((b) => (b.kind === 'paragraph' ? b.lines.flatMap((l) => l.line.fragments) : []));
    const body = texts(page?.blocks ?? []);
    return {
      body: body.map((f) => f.text).join(''),
      header: texts(page?.header?.blocks ?? [])
        .map((f) => f.text)
        .join(''),
      colors: body.map((f) => f.style.color),
    };
  }

  it('默认最终状态：被删的字不占位；换到「所有标记」正文与页眉一起重排，只派发 layout:done', async () => {
    const d = await UltimateWord.load(revisedDocx());
    expect(d.revisionView).toBe('final');
    expect(painted(d)).toMatchObject({ body: '原文新增结尾', header: '新眉' });
    expect(d.find('删掉')).toHaveLength(0);

    const events: string[] = [];
    d.on('layout:done', () => events.push('layout'));
    d.on('document:change', () => events.push('change'));
    d.setRevisionView('markup');
    expect(events).toEqual(['layout']);
    expect(d.canUndo).toBe(false);
    const markup = painted(d);
    expect(markup).toMatchObject({ body: '原文删掉新增结尾', header: '新眉旧眉' });
    // 作者按第一次出现配色：张三（删）第一色、李四（增）第二色
    expect(new Set(markup.colors).size).toBe(3);
    expect(d.find('删掉')).toHaveLength(1);

    d.setRevisionView('original');
    expect(painted(d)).toMatchObject({ body: '原文删掉结尾', header: '旧眉' });
  });

  it('LoadOptions.revisions 指定初始视图；修订列表与视图无关；视图下照常编辑', async () => {
    const d = await UltimateWord.load(revisedDocx(), { revisions: 'markup' });
    expect(d.revisionView).toBe('markup');
    expect(painted(d).body).toBe('原文删掉新增结尾');
    expect(d.revisions().map(({ kind, author, text }) => ({ kind, author, text }))).toEqual([
      { kind: 'delete', author: '张三', text: '删掉' },
      { kind: 'insert', author: '李四', text: '新增' },
    ]);
    // 编辑后重排仍按当前视图（修订的配色不因编辑漂移）
    const before = painted(d).colors;
    const [hit] = d.find('结尾');
    if (hit === undefined) throw new Error('找不到「结尾」');
    d.tx((t) => void t.insertText(hit.start, '甲'));
    expect(painted(d).body).toBe('原文删掉新增甲结尾');
    expect(painted(d).colors.slice(0, 3)).toEqual(before.slice(0, 3));
  });

  it('替换跳过首字落在被删除修订里的命中，不让整批回滚', async () => {
    const d = await UltimateWord.load(revisedDocx(), { revisions: 'markup' });
    expect(d.replaceAll(/删掉|结尾/, '改')).toEqual({ replaced: 1, skipped: 1 });
    expect(painted(d).body).toBe('原文删掉新增改');
  });
});
