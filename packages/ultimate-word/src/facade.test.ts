/**
 * 门面的判据不是真值残差 —— 字摆在哪由底下已标定完的几层决定，这里只验「接线没接错」：
 * 四种 `LoadSource` 都通到同一份布局、`find` 自动带上域求值结果、`rangeOf` 认 id 也认节点、
 * 随库度量包在模块加载时就已经在注册表里。跑在纯 Node 上，`mount()` 的那一半在
 * `apps/playground/tests/facade.html`。
 */
import { readFileSync } from 'node:fs';
import { FALLBACK_METRICS, FontRegistry } from '@uw/fonts';
import type { DocRange } from '@uw/model';
import { unzip, zip } from '@uw/ooxml';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import type { BindingValue, UwDocument } from './index.ts';
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
  function template(extra = ''): Uint8Array {
    const enc = new TextEncoder();
    const body =
      `<w:p><w:r><w:t>申请人：</w:t></w:r>${control('applicant', '单击输入', '<w:showingPlcHdr/><w:text/>')}</w:p>` +
      `<w:p><w:r><w:t>日期：</w:t></w:r>${control('date', '某日')}<w:r><w:t>，签名：</w:t></w:r>${control('applicant', '单击输入', '<w:showingPlcHdr/><w:text/>')}</w:p>` +
      `<w:p>${control('level', '一般', '<w:dropDownList><w:listItem w:displayText="紧急" w:value="1"/><w:listItem w:displayText="一般" w:value="2"/></w:dropDownList>')}${control('agree', '☐', '<w14:checkbox/>')}</w:p>${extra}`;
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

  it('setMany 后续字段类型错误时不覆盖旧草稿，也不留下部分新值', async () => {
    const d = await UltimateWord.load(template());
    const pending = d.bindings.set('applicant', '原草稿').set('date', '旧日期').pending;
    const before = [...pending];
    expect(() => d.bindings.setMany({ applicant: '意外覆盖', level: '1', agree: '错误类型' })).toThrow(
      /布尔值/,
    );
    expect([...pending]).toEqual(before);
    d.bindings.apply();
    expect(
      d.bindings
        .list()
        .filter((b) => b.key === 'applicant')
        .map((b) => b.value),
    ).toEqual(['原草稿', '原草稿']);
    expect(d.bindings.list().find((b) => b.key === 'level')?.value).toBe('一般');
  });

  it.each([
    ['文字类型', '<w:text/>', true],
    ['无效日期', '<w:date><w:dateFormat w:val="yyyy-MM-dd"/></w:date>', { date: '2026-02-29' }],
    ['日期缺格式', '<w:date/>', { date: '2026-10-23' }],
    ['图片控件', '<w:picture/>', '文字'],
    ['同名混合类型', '<w:date><w:dateFormat w:val="yyyy-MM-dd"/></w:date>', { date: '2026-10-23' }],
  ] as const)('setMany 遇到%s时保留旧队列、模型、历史与导出', async (kind, props, value) => {
    const extra = `<w:p>${control('invalid', '原值', props)}${kind === '同名混合类型' ? control('invalid', '文字', '<w:text/>') : ''}</w:p>`;
    const d = await UltimateWord.load(template(extra));
    d.bindings.set('applicant', '历史值').apply();
    d.undo();
    const before = d.bindings.list();
    const original = new Uint8Array(await (await d.toDocx()).arrayBuffer());
    const pending = d.bindings.set('applicant', '旧草稿').set('date', '旧日期').pending;
    const staged = [...pending];
    const events: string[] = [];
    d.on('document:change', () => events.push('change'));
    d.on('layout:done', () => events.push('layout'));
    expect(() => d.bindings.setMany({ applicant: '意外覆盖', agree: true, invalid: value })).toThrow();
    expect(d.bindings.pending).toBe(pending);
    expect([...pending]).toEqual(staged);
    expect(d.bindings.list()).toEqual(before);
    expect(events).toEqual([]);
    expect(d.canUndo).toBe(false);
    expect(d.canRedo).toBe(true);
    expect(new Uint8Array(await (await d.toDocx()).arrayBuffer())).toEqual(original);
    d.redo();
    expect(d.bindings.list().find((b) => b.key === 'applicant')?.value).toBe('历史值');
    expect([...pending]).toEqual(staged);
  });

  it('setMany 拒绝运行时无效值时不暂存前面的有效字段', async () => {
    const d = await UltimateWord.load(template());
    expect(() => d.bindings.setMany({ applicant: '张三', agree: null as unknown as BindingValue })).toThrow(
      TypeError,
    );
    expect(d.bindings.pending.size).toBe(0);
    expect(d.bindings.apply()).toBeUndefined();
    expect(d.canUndo).toBe(false);
  });

  it('setMany 失败保留原日期对象，已准备的新日期也不会冻结调用方对象', async () => {
    const d = await UltimateWord.load(
      template(
        `<w:p>${control('issued', '某日', '<w:date><w:dateFormat w:val="yyyy-MM-dd"/></w:date>')}</w:p>`,
      ),
    );
    const date = { date: '2026-10-23' };
    d.bindings.set('issued', { date: '2026-10-22' });
    const previous = d.bindings.pending.get('issued');
    expect(() => d.bindings.setMany({ applicant: '张三', issued: date, agree: '错误类型' })).toThrow();
    expect([...d.bindings.pending]).toEqual([['issued', { date: '2026-10-22' }]]);
    expect(d.bindings.pending.get('issued')).toBe(previous);
    expect(Object.isFrozen(date)).toBe(false);
    date.date = '2026-10-24';
    d.bindings.apply();
    expect(d.bindings.list().find((b) => b.key === 'issued')?.fullDate).toBe('2026-10-22T00:00:00Z');
  });

  it('setMany 成功按输入顺序合入原队列，日期复制冻结，一次提交并导出回读', async () => {
    const d = await UltimateWord.load(
      template(
        `<w:p>${control('issued', '某日', '<w:date><w:dateFormat w:val="yyyy-MM-dd"/></w:date>')}</w:p>`,
      ),
    );
    const before = d.bindings.list();
    const pending = d.bindings.set('date', '保留').set('applicant', '旧草稿').pending;
    const date = { date: '2026-10-23' };
    expect(d.bindings.setMany({ applicant: '张三', issued: date, agree: true, unused: '跳过' })).toEqual({
      skipped: ['unused'],
    });
    expect([...pending]).toEqual([
      ['date', '保留'],
      ['applicant', '张三'],
      ['issued', { date: '2026-10-23' }],
      ['agree', true],
    ]);
    expect(Object.isFrozen(pending.get('issued'))).toBe(true);
    date.date = '2026-10-24';
    expect(d.bindings.list()).toEqual(before);
    const events: string[] = [];
    d.on('document:change', () => events.push('change'));
    d.on('layout:done', () => events.push('layout'));
    expect(d.bindings.apply()).toBeDefined();
    expect(events).toEqual(['change', 'layout']);
    expect(pending.size).toBe(0);
    const after = d.bindings.list();
    expect(after.filter((b) => b.key === 'applicant').map((b) => b.value)).toEqual(['张三', '张三']);
    expect(after.find((b) => b.key === 'issued')?.fullDate).toBe('2026-10-23T00:00:00Z');
    const again = await UltimateWord.load(await d.toDocx());
    expect(again.bindings.list()).toEqual(after);
    d.undo();
    expect(d.bindings.list()).toEqual(before);
    expect(d.canUndo).toBe(false);
    d.redo();
    expect(d.bindings.list()).toEqual(after);
  });

  it('setMany 空对象和未知字段只返回 skipped，不校验未知值也不改变旧草稿', async () => {
    const d = await UltimateWord.load(template());
    d.bindings.set('applicant', '旧草稿');
    const before = [...d.bindings.pending];
    expect(d.bindings.setMany({})).toEqual({ skipped: [] });
    expect(
      d.bindings.setMany({ unknown: null as unknown as BindingValue, unused: { date: '无效' } }),
    ).toEqual({ skipped: ['unknown', 'unused'] });
    expect([...d.bindings.pending]).toEqual(before);
    expect(d.canUndo).toBe(false);
  });

  it.each(['document:change', 'layout:done'] as const)(
    '%s 中批量暂存失败也只保留先前待填值',
    async (event) => {
      const d = await UltimateWord.load(template());
      let error: unknown;
      const sub = d.on(event, () => {
        d.bindings.set('agree', true);
        try {
          d.bindings.setMany({ applicant: '意外覆盖', date: true });
        } catch (e) {
          error = e;
        }
      });
      d.bindings.set('applicant', '张三').apply();
      sub.dispose();
      expect(error).toBeInstanceOf(TypeError);
      expect([...d.bindings.pending]).toEqual([['agree', true]]);
      d.bindings.apply();
      expect(d.bindings.list().find((b) => b.key === 'applicant')?.value).toBe('张三');
      expect(d.bindings.list().find((b) => b.key === 'agree')?.checked).toBe(true);
      d.undo();
      expect(d.bindings.list().find((b) => b.key === 'applicant')?.value).toBe('张三');
      expect(d.bindings.list().find((b) => b.key === 'agree')?.checked).toBe(false);
    },
  );

  it.each(['document:change', 'layout:done'] as const)(
    '%s 回调暂存的下一批值不会被本次提交清空',
    async (event) => {
      const d = await UltimateWord.load(template());
      const pending = d.bindings.pending;
      let pendingInEvent = -1;
      const sub = d.on(event, () => {
        pendingInEvent = pending.size;
        d.bindings.setMany({ applicant: '李四', agree: true });
      });
      d.bindings.set('applicant', '张三').apply();
      sub.dispose();
      expect(pendingInEvent).toBe(0);
      expect([...pending]).toEqual([
        ['applicant', '李四'],
        ['agree', true],
      ]);
      expect(
        d.bindings
          .list()
          .filter((b) => b.key === 'applicant')
          .map((b) => b.value),
      ).toEqual(['张三', '张三']);
      expect(d.bindings.list().find((b) => b.key === 'agree')?.checked).toBe(false);
      expect(d.bindings.apply()).toBeDefined();
      expect(pending.size).toBe(0);
      const after = d.bindings.list();
      expect(after.filter((b) => b.key === 'applicant').map((b) => b.value)).toEqual(['李四', '李四']);
      expect(after.find((b) => b.key === 'agree')?.checked).toBe(true);
      const again = await UltimateWord.load(await d.toDocx());
      expect(again.bindings.list().map((b) => [b.key, b.value])).toEqual(after.map((b) => [b.key, b.value]));
      d.undo();
      expect(
        d.bindings
          .list()
          .filter((b) => b.key === 'applicant')
          .map((b) => b.value),
      ).toEqual(['张三', '张三']);
      expect(d.bindings.list().find((b) => b.key === 'agree')?.checked).toBe(false);
      d.undo();
      expect(d.canUndo).toBe(false);
      expect(d.bindings.list().find((b) => b.key === 'applicant')?.placeholder).toBe(true);
      d.redo();
      d.redo();
      expect(d.bindings.list()).toEqual(after);
    },
  );

  describe.each(['document:change', 'layout:done'] as const)('填写事件中的待填队列：%s', (event) => {
    it('重新暂存相同值仍属于下一批，空操作保留重做历史', async () => {
      const d = await UltimateWord.load(template());
      let changes = 0;
      const count = d.on('document:change', () => changes++);
      const sub = d.on(event, () => d.bindings.set('applicant', '张三'));
      d.bindings.set('applicant', '张三').apply();
      sub.dispose();
      count.dispose();
      expect([...d.bindings.pending]).toEqual([['applicant', '张三']]);
      d.bindings.set('applicant', '李四').apply();
      d.undo();
      d.bindings.set('applicant', '张三');
      expect(d.bindings.apply()).toBeUndefined();
      expect(d.bindings.pending.size).toBe(0);
      expect(changes).toBe(1);
      expect(d.canRedo).toBe(true);
    });

    it('没有新值时在回调内再次 apply 为空操作，不重复提交本批', async () => {
      const d = await UltimateWord.load(template());
      const results: unknown[] = [];
      let calls = 0;
      const sub = d.on(event, () => {
        if (++calls === 1) results.push(d.bindings.apply());
      });
      d.bindings.set('applicant', '张三').apply();
      sub.dispose();
      expect(results).toEqual([undefined]);
      expect(calls).toBe(1);
      expect(d.bindings.pending.size).toBe(0);
      d.undo();
      expect(d.canUndo).toBe(false);
    });

    it('回调内可以立即提交下一批，各自独立撤销且不重复旧字段', async () => {
      const d = await UltimateWord.load(template());
      const snapshots: string[][] = [];
      const sub = d.on(event, () => {
        snapshots.push([...d.bindings.pending.keys()]);
        if (snapshots.length === 1) d.bindings.set('agree', true).apply();
      });
      d.bindings.set('applicant', '张三').apply();
      sub.dispose();
      expect(snapshots).toEqual([[], []]);
      expect(d.bindings.pending.size).toBe(0);
      expect(d.bindings.list().find((b) => b.key === 'agree')?.checked).toBe(true);
      d.undo();
      expect(d.bindings.list().find((b) => b.key === 'agree')?.checked).toBe(false);
      expect(d.bindings.list().find((b) => b.key === 'applicant')?.value).toBe('张三');
      d.undo();
      expect(d.bindings.list().find((b) => b.key === 'applicant')?.placeholder).toBe(true);
      expect(d.canUndo).toBe(false);
      d.redo();
      d.redo();
      const again = await UltimateWord.load(await d.toDocx());
      expect(again.bindings.list().find((b) => b.key === 'agree')?.checked).toBe(true);
    });

    it('回调内下一批失败时保留整批待填值，已成功的外层提交不清掉它', async () => {
      const d = await UltimateWord.load(template());
      let error: unknown;
      const sub = d.on(event, () => {
        d.bindings.setMany({ agree: true, level: '不存在的选项' });
        try {
          d.bindings.apply();
        } catch (e) {
          error = e;
        }
      });
      expect(d.bindings.set('applicant', '张三').apply()).toBeDefined();
      sub.dispose();
      expect(error).toBeInstanceOf(Error);
      expect([...d.bindings.pending]).toEqual([
        ['agree', true],
        ['level', '不存在的选项'],
      ]);
      expect(d.bindings.list().find((b) => b.key === 'agree')?.checked).toBe(false);
      expect(d.bindings.list().find((b) => b.key === 'applicant')?.value).toBe('张三');
      d.bindings.set('level', '紧急').apply();
      expect(d.bindings.pending.size).toBe(0);
      expect(d.bindings.list().find((b) => b.key === 'agree')?.checked).toBe(true);
      d.undo();
      expect(d.bindings.list().find((b) => b.key === 'applicant')?.value).toBe('张三');
      expect(d.bindings.list().find((b) => b.key === 'agree')?.checked).toBe(false);
    });

    it('discard 后重新暂存日期对象仍保留复制的下一批值', async () => {
      const d = await UltimateWord.load(
        template(
          `<w:p>${control('issued', '某日', '<w:date><w:dateFormat w:val="yyyy-MM-dd"/></w:date>')}</w:p>`,
        ),
      );
      const date = { date: '2026-10-23' };
      const sub = d.on(event, () => {
        d.bindings.set('applicant', '丢弃');
        d.bindings.discard();
        d.bindings.set('issued', date);
        date.date = '2026-10-24';
      });
      d.bindings.set('applicant', '张三').apply();
      sub.dispose();
      expect([...d.bindings.pending]).toEqual([['issued', { date: '2026-10-23' }]]);
      expect(Object.isFrozen(d.bindings.pending.get('issued'))).toBe(true);
      d.bindings.apply();
      expect(d.bindings.list().find((b) => b.key === 'issued')?.fullDate).toBe('2026-10-23T00:00:00Z');
    });
  });

  it('事务回调内不允许嵌套 apply，拒绝后恢复原待填队列与历史', async () => {
    const d = await UltimateWord.load(template());
    const before = d.bindings.list();
    const pending = d.bindings.set('applicant', '张三').pending;
    expect(() => d.tx(() => void d.bindings.apply())).toThrow();
    expect(d.bindings.list()).toEqual(before);
    expect([...pending]).toEqual([['applicant', '张三']]);
    expect(d.canUndo).toBe(false);
    expect(d.bindings.apply()).toBeDefined();
    expect(pending.size).toBe(0);
  });

  it('命令结束后的预排版失败也恢复整批待填值，重试只产生一个撤销单元', async () => {
    const fonts = new FontRegistry();
    const d = await UltimateWord.load(template(), { fonts });
    const before = d.bindings.list();
    const layout = d.layout;
    const original = new Uint8Array(await (await d.toDocx()).arrayBuffer());
    d.bindings.setMany({ applicant: '张三', agree: true });
    const pending = [...d.bindings.pending];
    // 注册新度量使重排不能复用加载时的缓存，故障发生在命令成功之后的提交校验。
    fonts.register('预排版回滚', { kind: 'file', metrics: FALLBACK_METRICS, advance: () => 1000 });
    const failure = new Error('预排版失败');
    const resolve = vi.spyOn(fonts, 'resolve').mockImplementation(() => {
      throw failure;
    });
    try {
      expect(() => d.bindings.apply()).toThrow(failure);
    } finally {
      resolve.mockRestore();
    }
    expect(d.bindings.list()).toEqual(before);
    expect(d.layout).toBe(layout);
    expect([...d.bindings.pending]).toEqual(pending);
    expect(d.canUndo).toBe(false);
    expect(new Uint8Array(await (await d.toDocx()).arrayBuffer())).toEqual(original);
    expect(d.bindings.apply()).toBeDefined();
    expect(d.bindings.pending.size).toBe(0);
    d.undo();
    expect(d.bindings.list()).toEqual(before);
    expect(d.canUndo).toBe(false);
  });

  it.each([
    ['date', 'paragraph'],
    ['checkbox', 'paragraph'],
    ['date', 'row'],
    ['checkbox', 'row'],
    ['date', 'outerRow'],
    ['checkbox', 'outerRow'],
  ] as const)('%s 在 %s 修订中拒绝填写，三种视图下批量回滚、历史与原文件一致', async (type, scope) => {
    const pr =
      type === 'date'
        ? '<w:date w:fullDate="2026-10-13T00:00:00Z"><w:dateFormat w:val="yyyy"/></w:date>'
        : '<w14:checkbox/>';
    for (const revision of ['ins', 'del'])
      for (const revisions of ['final', 'original', 'markup'] as const) {
        const mark = `<w:${revision} w:id="7" w:author="甲"/>`;
        const paragraph = `<w:p>${scope === 'paragraph' ? `<w:pPr><w:rPr>${mark}</w:rPr></w:pPr>` : ''}${control('revised', type === 'date' ? '2026' : '☐', `<w:dataBinding w:xpath="/field"/>${pr}`)}</w:p>`;
        const table = (content: string, revised = true) =>
          `<w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr>${revised ? `<w:trPr>${mark}</w:trPr>` : ''}<w:tc>${content}</w:tc></w:tr></w:tbl>`;
        const bytes = template(
          `${scope === 'paragraph' ? paragraph : table(scope === 'outerRow' ? `${table(paragraph, false)}<w:p/>` : paragraph)}<w:p/>`,
        );
        const d = await UltimateWord.load(bytes, { revisions });
        let changes = 0;
        d.on('document:change', () => changes++);
        d.on('layout:done', () => changes++);
        const initial = d.bindings.list();
        const field = initial.find((b) => b.key === 'revised');
        if (field === undefined) throw new Error('样本缺少控件');
        const value = type === 'date' ? { date: '2026-10-14' } : true;
        d.bindings.setMany({ applicant: '张三', revised: value });
        expect(() => d.bindings.apply()).toThrow(/无修订/);
        expect(d.bindings.list()).toEqual(initial);
        expect(d.bindings.pending.get('revised')).toEqual(value);
        expect(changes).toBe(0);
        expect(d.canUndo).toBe(false);
        expect(() =>
          d.tx((t) => {
            if (type === 'date') t.setContentControlDate(field.id, '2026-10-13');
            else t.setContentControlChecked(field.id, false);
          }),
        ).toThrow(/无修订/);
        const out = new Uint8Array(await (await d.toDocx()).arrayBuffer());
        const sourceParts = unzip(bytes);
        for (const [name, content] of unzip(out)) expect(content).toEqual(sourceParts.get(name));
        const again = await UltimateWord.load(out, { revisions });
        expect(again.bindings.list()).toEqual(initial);
      }
  });

  it.each(['<w:p/><w:p/>', '<w:p/><w:p><w:r><w:t>原值</w:t></w:r></w:p><w:p/>'])(
    '块级控件范围包含边界空段，填写与导出回读一致：%s',
    async (content) => {
      const d = await UltimateWord.load(
        template(
          `<w:sdt><w:sdtPr><w:tag w:val="outer"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`,
        ),
      );
      const before = d.bindings.list().find((b) => b.key === 'outer');
      if (before?.range === undefined) throw new Error('样本缺少范围');
      const paragraphs = d
        .query('paragraph')
        .filter((p) => p.kind === 'paragraph' && p.contentControl === before.id);
      expect(before.range.start.nodeId).toBe(paragraphs[0]?.id);
      expect(before.range.end.nodeId).toBe(paragraphs.at(-1)?.id);
      let changes = 0;
      d.on('document:change', () => changes++);
      d.bindings.set('outer', '新值').apply();
      expect(changes).toBe(1);
      expect(d.bindings.list().find((b) => b.key === 'outer')?.value).toBe('新值');
      const again = await UltimateWord.load(await d.toDocx());
      expect(again.bindings.list().find((b) => b.key === 'outer')?.value).toBe('新值');
      d.undo();
      expect(d.bindings.list().find((b) => b.key === 'outer')).toEqual(before);
      d.redo();
      expect(d.bindings.list().find((b) => b.key === 'outer')?.value).toBe('新值');
    },
  );

  it.each([false, true])('占位控件填空后再输入继承控件格式，事件与历史一致：块级 %s', async (block) => {
    const run = '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/><w:i/></w:rPr><w:t>请输入</w:t></w:r>';
    const target = `<w:sdt><w:sdtPr><w:rPr><w:b/></w:rPr><w:tag w:val="emptyPlaceholder"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent>${block ? `<w:p>${run}</w:p>` : run}</w:sdtContent></w:sdt>`;
    const d = await UltimateWord.load(template(block ? target : `<w:p>${target}</w:p>`));
    const original = await d.toDocx();
    let changes = 0;
    d.on('document:change', () => changes++);
    d.bindings.set('emptyPlaceholder', '').apply();
    const empty = d.bindings.list().find((b) => b.key === 'emptyPlaceholder');
    if (empty?.range === undefined) throw new Error('样本缺少范围');
    const range = empty.range;
    expect(changes).toBe(1);
    d.tx((t) => void t.insertText(range.end, '张三'));
    expect(d.query('run').find((r) => r.id === empty.range?.start.nodeId)?.props).toEqual({ bold: true });
    const again = await UltimateWord.load(await d.toDocx());
    expect(again.find('张三')).toHaveLength(1);
    expect(again.query('run').find((r) => r.id === again.find('张三')[0]?.start.nodeId)?.props).toEqual({
      bold: true,
    });
    d.undo();
    d.bindings.set('emptyPlaceholder', '').apply();
    expect(changes).toBe(3);
    expect(d.canRedo).toBe(true);
    d.undo();
    expect(await d.toDocx()).toEqual(original);
    d.redo();
    d.redo();
    expect(d.find('张三')).toHaveLength(1);
  });

  it('覆盖首个子控件只留下目标，保留相邻字段与撤销导出', async () => {
    const inner = control('firstChild', '原值', '<w:text/>');
    const target = `<w:sdt><w:sdtPr><w:tag w:val="outer"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
    const d = await UltimateWord.load(template(`<w:p>${target}${control('neighbor', '旁边')}</w:p>`));
    const original = await d.toDocx();
    const start = d.find('原值')[0]?.start;
    let events = 0;
    d.on('document:change', () => events++);
    d.bindings.set('outer', '甲\n乙').apply();
    expect(events).toBe(1);
    expect(d.bindings.list().find((b) => b.key === 'firstChild')).toBeUndefined();
    expect(d.bindings.list().find((b) => b.key === 'outer')?.value).toBe('甲\n乙');
    expect(d.bindings.list().find((b) => b.key === 'neighbor')?.value).toBe('旁边');
    expect(d.find('甲')[0]?.start).toEqual(start);
    const again = await UltimateWord.load(await d.toDocx());
    expect(again.bindings.list().map((b) => [b.key, b.value])).toEqual(
      d.bindings.list().map((b) => [b.key, b.value]),
    );
    d.undo();
    expect(await (await d.toDocx()).arrayBuffer()).toEqual(await original.arrayBuffer());
    d.redo();
    expect(d.bindings.list().find((b) => b.key === 'firstChild')).toBeUndefined();
  });

  function nestedTemplate(scope: string, pr = '', sameName = false): Uint8Array {
    const wrap = (tag: string, content: string, props = '') =>
      `<w:sdt><w:sdtPr><w:tag w:val="${sameName ? 'nested' : tag}"/>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
    const inner = wrap('inner', '<w:r><w:rPr><w:b/></w:rPr><w:t>原值</w:t></w:r>', pr);
    const middle = wrap('middle', scope === 'nestedBlock' ? `<w:p>${inner}</w:p>` : inner);
    const outer = wrap(
      'outer',
      scope === 'mixed'
        ? `<w:p>${middle}</w:p>`
        : scope === 'inlinePrefix'
          ? `<w:r><w:t>前缀</w:t></w:r>${middle}`
          : middle,
    );
    return template(
      `${scope === 'nestedBlock' || scope === 'mixed' ? outer : `<w:p>${outer}</w:p>`}<w:p>${control('neighbor', '旁边')}</w:p>`,
    );
  }

  describe.each(['inline', 'inlinePrefix', 'nestedBlock', 'mixed'])('嵌套批量填写：%s', (scope) => {
    it.each([false, true])('暂存字段被移除时拒绝提交，混合填写 %s，恢复后可重试', async (mixed) => {
      const d = await UltimateWord.load(nestedTemplate(scope));
      const initial = d.bindings.list();
      const outer = initial.find((b) => b.key === 'outer');
      if (outer === undefined) throw new Error('样本缺少外层控件');
      d.bindings.set('inner', '子更新');
      if (mixed) d.bindings.set('neighbor', '邻居更新');
      const pending = [...d.bindings.pending];
      d.tx((t) => {
        t.fillContentControl(outer.id, '覆盖');
      });
      const applicant = d.bindings.list().find((b) => b.key === 'applicant');
      if (applicant === undefined) throw new Error('样本缺少申请人控件');
      d.tx((t) => {
        t.fillContentControl(applicant.id, '历史值');
      });
      d.undo();
      const before = d.bindings.list();
      const original = new Uint8Array(await (await d.toDocx()).arrayBuffer());
      let changes = 0;
      let layouts = 0;
      d.on('document:change', () => changes++);
      d.on('layout:done', () => layouts++);
      expect(() => d.bindings.apply()).toThrow(/文档里没有坑位：inner/);
      expect(changes).toBe(0);
      expect(layouts).toBe(0);
      expect(d.bindings.list()).toEqual(before);
      expect([...d.bindings.pending]).toEqual(pending);
      expect(new Uint8Array(await (await d.toDocx()).arrayBuffer())).toEqual(original);
      expect(d.canRedo).toBe(true);
      d.redo();
      expect(d.bindings.list().find((b) => b.key === 'applicant')?.value).toBe('历史值');
      d.undo();
      d.undo();
      expect(d.bindings.list()).toEqual(initial);
      expect([...d.bindings.pending]).toEqual(pending);
      changes = 0;
      layouts = 0;
      expect(d.bindings.apply()).toBeDefined();
      expect(changes).toBe(1);
      expect(layouts).toBe(1);
      expect(d.bindings.pending.size).toBe(0);
      const after = d.bindings.list();
      expect(after.find((b) => b.key === 'inner')?.value).toBe('子更新');
      expect(after.find((b) => b.key === 'neighbor')?.value).toBe(mixed ? '邻居更新' : '旁边');
      d.undo();
      expect(d.bindings.list()).toEqual(initial);
      d.redo();
      expect(d.bindings.list()).toEqual(after);
      const again = await UltimateWord.load(await d.toDocx());
      expect(again.bindings.list().map((b) => [b.key, b.value])).toEqual(after.map((b) => [b.key, b.value]));
    });

    it.each(['最终值', '原值', ''])('从内到外填写 %j，一个撤销单元且导出一致', async (value) => {
      const d = await UltimateWord.load(nestedTemplate(scope));
      const before = d.bindings.list();
      const original = new Uint8Array(await (await d.toDocx()).arrayBuffer());
      let changes = 0;
      let layouts = 0;
      d.on('document:change', () => changes++);
      d.on('layout:done', () => layouts++);
      d.bindings.setMany({ outer: value, inner: '子更新', middle: '中层更新', neighbor: '邻居更新' });
      expect(d.bindings.apply()).toBeDefined();
      expect(changes).toBe(1);
      expect(layouts).toBe(1);
      expect(d.bindings.pending.size).toBe(0);
      const after = d.bindings.list();
      expect(after.find((b) => b.key === 'outer')?.value).toBe(value);
      expect(after.some((b) => b.key === 'inner' || b.key === 'middle')).toBe(false);
      expect(after.find((b) => b.key === 'neighbor')?.value).toBe('邻居更新');
      const again = await UltimateWord.load(await d.toDocx());
      expect(again.bindings.list().map((b) => [b.key, b.value])).toEqual(after.map((b) => [b.key, b.value]));
      d.undo();
      expect(d.bindings.list()).toEqual(before);
      expect(d.canUndo).toBe(false);
      const sourceParts = unzip(original);
      for (const [name, content] of unzip(new Uint8Array(await (await d.toDocx()).arrayBuffer())))
        expect(content).toEqual(sourceParts.get(name));
      d.redo();
      expect(d.bindings.list()).toEqual(after);
    });

    it('同名祖孙字段与未提交的中间层均遵守父子顺序', async () => {
      for (const sameName of [false, true]) {
        const d = await UltimateWord.load(nestedTemplate(scope, '', sameName));
        if (sameName) d.bindings.set('nested', '同名新值');
        else d.bindings.setMany({ inner: '子更新', outer: '同名新值' });
        d.bindings.apply();
        expect(
          d.bindings.list().filter((b) => ['nested', 'outer', 'middle', 'inner'].includes(b.key)),
        ).toEqual([expect.objectContaining({ key: sameName ? 'nested' : 'outer', value: '同名新值' })]);
      }
    });

    it.each(['sdtLocked', 'contentLocked'])('子控件 %s 仍阻止整批填写，保留待填值与重做', async (lock) => {
      const d = await UltimateWord.load(nestedTemplate(scope, `<w:lock w:val="${lock}"/>`));
      d.bindings.set('applicant', '张三').apply();
      d.undo();
      const before = d.bindings.list();
      let changes = 0;
      d.on('document:change', () => changes++);
      d.bindings.setMany({ inner: '子更新', outer: '覆盖', neighbor: '邻居更新' });
      expect(() => d.bindings.apply()).toThrow(/锁定/);
      expect(d.bindings.list()).toEqual(before);
      expect([...d.bindings.pending]).toEqual([
        ['inner', '子更新'],
        ['outer', '覆盖'],
        ['neighbor', '邻居更新'],
      ]);
      expect(changes).toBe(0);
      expect(d.canRedo).toBe(true);
      d.redo();
      expect(d.bindings.list().find((b) => b.key === 'applicant')?.value).toBe('张三');
    });
  });

  it('暂存字段只移除部分同名控件时，仍按当前名字填写剩余控件', async () => {
    const d = await UltimateWord.load(nestedTemplate('inlinePrefix', '', true));
    const outerId = d.bindings.list().find((b) => b.key === 'nested' && b.value === '前缀原值')?.id;
    if (outerId === undefined) throw new Error('样本缺少外层控件');
    d.bindings.set('nested', '暂存值');
    d.tx((t) => {
      t.fillContentControl(outerId, '覆盖');
    });
    expect(d.bindings.list().filter((b) => b.key === 'nested')).toHaveLength(1);
    expect(d.bindings.apply()).toBeDefined();
    expect(d.bindings.list().filter((b) => b.key === 'nested')).toEqual([
      expect.objectContaining({ value: '暂存值' }),
    ]);
    expect(d.bindings.pending.size).toBe(0);
  });

  it('丢弃已移除字段的待填值后，可以提交新一批数据', async () => {
    const d = await UltimateWord.load(nestedTemplate('inline'));
    const outer = d.bindings.list().find((b) => b.key === 'outer');
    if (outer === undefined) throw new Error('样本缺少控件');
    d.bindings.set('inner', '子更新');
    d.tx((t) => {
      t.fillContentControl(outer.id, '覆盖');
    });
    expect(() => d.bindings.apply()).toThrow(RangeError);
    d.bindings.discard();
    expect(d.bindings.apply()).toBeUndefined();
    expect(d.bindings.setMany({ inner: '已移除', neighbor: '新一批' })).toEqual({ skipped: ['inner'] });
    expect(d.bindings.apply()).toBeDefined();
    expect(d.bindings.list().find((b) => b.key === 'neighbor')?.value).toBe('新一批');
  });

  it('无前缀的锁定子外壳阻止外层覆盖，失败保留待填值与重做', async () => {
    const inner = control('firstChild', '原值', '<w:lock w:val="sdtLocked"/>');
    const d = await UltimateWord.load(
      template(
        `<w:p><w:sdt><w:sdtPr><w:tag w:val="outer"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt></w:p>`,
      ),
    );
    d.bindings.set('applicant', '张三').apply();
    d.undo();
    let events = 0;
    d.on('document:change', () => events++);
    d.bindings.set('outer', '原值');
    expect(() => d.bindings.apply()).toThrow(/锁定了外壳/);
    expect(events).toBe(0);
    expect(d.bindings.pending.get('outer')).toBe('原值');
    expect(d.bindings.list().find((b) => b.key === 'firstChild')?.value).toBe('原值');
    expect(d.canRedo).toBe(true);
    d.redo();
    expect(d.bindings.list().find((b) => b.key === 'applicant')?.value).toBe('张三');
  });

  it('覆盖外层不能移除锁定子外壳，失败不派发事件、保留待填值与历史', async () => {
    const inner = control('innerShell', '子控件', '<w:lock w:val="sdtLocked"/>');
    const d = await UltimateWord.load(
      template(
        `<w:sdt><w:sdtPr><w:tag w:val="outerShell"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>前缀</w:t></w:r>${inner}</w:p></w:sdtContent></w:sdt>`,
      ),
    );
    d.bindings.set('applicant', '张三').apply();
    let events = 0;
    d.on('document:change', () => events++);
    d.bindings.set('outerShell', '覆盖');
    expect(() => d.bindings.apply()).toThrow(/锁定了外壳/);
    expect(d.bindings.pending.get('outerShell')).toBe('覆盖');
    expect(events).toBe(0);
    expect(d.bindings.list().find((b) => b.key === 'innerShell')).toMatchObject({
      value: '子控件',
      locked: false,
    });
    expect(d.find('前缀子控件')).toHaveLength(1);
    d.undo();
    expect(d.find('张三')).toHaveLength(0);
    d.redo();
    const again = await UltimateWord.load(await d.toDocx());
    expect(again.bindings.list().find((b) => b.key === 'innerShell')?.value).toBe('子控件');
    again.bindings.set('outerShell', '再覆盖');
    expect(() => again.bindings.apply()).toThrow(/锁定了外壳/);
  });

  it('外壳内部删光可导出并重新填写，历史跳转和回写保留锁', async () => {
    const d = await UltimateWord.load(
      template(`<w:p>${control('shell', '壳内文字', '<w:lock w:val="sdtLocked"/>')}</w:p>`),
    );
    const shell = d.find('壳内文字')[0] as DocRange;
    d.tx((t) => void t.deleteRange(shell));
    const empty = await UltimateWord.load(await d.toDocx());
    expect(empty.bindings.list().find((b) => b.key === 'shell')).toMatchObject({ value: '', locked: false });
    empty.bindings.set('shell', '重新填写').apply();
    empty.undo();
    expect(empty.bindings.list().find((b) => b.key === 'shell')?.value).toBe('');
    empty.redo();
    const exported = await empty.toDocx();
    const again = await UltimateWord.load(exported);
    expect(again.bindings.list().find((b) => b.key === 'shell')?.value).toBe('重新填写');
    const xml = new TextDecoder().decode(
      unzip(new Uint8Array(await exported.arrayBuffer())).get('word/document.xml'),
    );
    expect(xml).toContain('<w:lock w:val="sdtLocked"');
  });

  it('普通文字命令不能绕过内容锁，失败不派发变更也不破坏已有撤销历史', async () => {
    const d = await UltimateWord.load(
      template(`<w:p>${control('locked', '锁中', '<w:lock w:val="contentLocked"/>')}</w:p>`),
    );
    d.bindings.set('applicant', '张三').apply();
    let events = 0;
    d.on('document:change', () => events++);
    const locked = d.find('锁中')[0] as DocRange;
    const applicant = d.find('张三')[0] as DocRange;
    expect(() =>
      d.tx((t) => {
        t.insertText(applicant.end, '先改');
        t.replaceText(locked, '篡改');
      }),
    ).toThrow(/锁定/);
    expect(events).toBe(0);
    expect(d.find('张三')).toHaveLength(2);
    expect(d.find('先改')).toHaveLength(0);
    expect(d.find('锁中')).toHaveLength(1);
    d.undo();
    expect(d.find('张三')).toHaveLength(0);
    d.redo();
    expect(d.find('张三')).toHaveLength(2);
  });

  it('全部替换跳过锁定命中，外壳锁仍可改，撤销与导出再加载保留锁', async () => {
    const d = await UltimateWord.load(
      template(
        `<w:p>${control('locked', '一般', '<w:lock w:val="sdtContentLocked"/>')}${control('shell', '一般', '<w:lock w:val="sdtLocked"/>')}</w:p>`,
      ),
    );
    let callbacks = 0;
    expect(
      d.replaceAll('一般', () => {
        callbacks++;
        return '紧急';
      }),
    ).toEqual({ replaced: 1, skipped: 2 });
    expect(callbacks).toBe(1);
    expect(d.find('紧急')).toHaveLength(1);
    expect(d.find('一般')).toHaveLength(2);
    d.undo();
    expect(d.find('一般')).toHaveLength(3);
    d.redo();
    const again = await UltimateWord.load(await d.toDocx());
    expect(again.bindings.list().find((b) => b.key === 'locked')).toMatchObject({
      value: '一般',
      locked: true,
    });
    expect(again.replaceAll('一般', '不改')).toEqual({ replaced: 0, skipped: 2 });
    expect(again.find('紧急')).toHaveLength(1);
  });

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

  it('列表提供完整选项与保存值，填写与点击所用元数据同撤销并可导出，返回选项不泄漏可变引用', async () => {
    const d = await UltimateWord.load(
      template(
        `<w:p>${control('other', '选择', '<w:comboBox w:lastValue="old"><w:listItem w:displayText="紧急" w:value="1"/></w:comboBox>')}</w:p>`,
      ),
    );
    const level = d.bindings.list().find((b) => b.key === 'level');
    expect(level?.options).toEqual(['紧急', '一般']);
    expect(level?.items).toEqual([
      { text: '紧急', value: '1' },
      { text: '一般', value: '2' },
    ]);
    expect(level?.lastValue).toBeUndefined();
    if (level?.items?.[0]) level.items[0].value = '篡改';
    expect(d.bindings.list().find((b) => b.key === 'level')?.items?.[0]?.value).toBe('1');
    d.bindings.setMany({ level: '1', other: '自由输入' });
    d.bindings.apply();
    expect(
      d.bindings
        .list()
        .filter((b) => b.items)
        .map((b) => [b.value, b.lastValue]),
    ).toEqual([
      ['紧急', '1'],
      ['自由输入', '自由输入'],
    ]);
    d.undo();
    expect(d.bindings.list().find((b) => b.key === 'level')?.lastValue).toBeUndefined();
    expect(d.bindings.list().find((b) => b.key === 'other')?.lastValue).toBe('old');
    d.redo();
    const again = await UltimateWord.load(await d.toDocx());
    expect(
      again.bindings
        .list()
        .filter((b) => b.items)
        .map((b) => [b.value, b.lastValue]),
    ).toEqual([
      ['紧急', '1'],
      ['自由输入', '自由输入'],
    ]);
    d.bindings.set('level', '紧急');
    expect(d.bindings.apply()).toBeUndefined();
  });

  it('空下拉框拒绝任意值，混合填写回滚保存值并保留 pending', async () => {
    const d = await UltimateWord.load(
      template(`<w:p>${control('emptyList', '选择', '<w:dropDownList/>')}</w:p>`),
    );
    d.bindings.setMany({ level: '1', emptyList: '自由输入' });
    expect(() => d.bindings.apply()).toThrow(/下拉框没有这个选项/);
    expect(d.bindings.list().find((b) => b.key === 'level')).toMatchObject({ value: '一般' });
    expect(d.bindings.list().find((b) => b.key === 'level')?.lastValue).toBeUndefined();
    expect(d.bindings.pending.get('level')).toBe('1');
    expect(d.bindings.pending.get('emptyList')).toBe('自由输入');
  });

  it('单行控件拒绝换行且不派发事件，正常填写可撤销并保留导出回读后的限制', async () => {
    const d = await UltimateWord.load(template());
    let events = 0;
    d.on('document:change', () => {
      events++;
    });
    const r = d.find('单击输入')[0] as DocRange;
    expect(() => d.tx((t) => void t.replaceText(r, '甲\n乙'))).toThrow(/单行纯文本控件不能换行/);
    expect(() => d.replaceAll('单击输入', '甲\n乙')).toThrow(/单行纯文本控件不能换行/);
    expect(events).toBe(0);
    expect(d.canUndo).toBe(false);
    expect(d.find('单击输入')).toHaveLength(2);
    d.bindings.set('applicant', '甲\n乙').apply();
    expect(events).toBe(1);
    expect(d.find('甲 乙')).toHaveLength(2);
    d.undo();
    expect(d.find('单击输入')).toHaveLength(2);
    d.redo();
    expect(d.find('甲 乙')).toHaveLength(2);
    const again = await UltimateWord.load(await d.toDocx());
    expect(
      again.bindings
        .list()
        .filter((b) => b.key === 'applicant')
        .every((b) => b.value === '甲 乙'),
    ).toBe(true);
    expect(() => again.tx((t) => void t.splitParagraph((again.find('甲 乙')[0] as DocRange).end))).toThrow(
      /单行纯文本控件不能换行/,
    );
    expect(again.replaceAll('甲 乙', '张三')).toEqual({ replaced: 2, skipped: 0 });
  });

  it('坑位给出独立 id，按选项序号填写只改当前同名控件并精确回写保存值', async () => {
    const pr =
      '<w:dropDownList w:lastValue="a"><w:listItem w:displayText="同名" w:value="a"/><w:listItem w:displayText="同名" w:value="b"/><w:listItem w:displayText="b" w:value="c"/></w:dropDownList>';
    const d = await UltimateWord.load(
      template(`<w:p>${control('duplicate', '同名', pr)}${control('duplicate', '同名', pr)}</w:p>`),
    );
    const list = d.bindings.list().filter((b) => b.key === 'duplicate');
    expect(list[0]?.id).not.toBe(list[1]?.id);
    const id = list[0]?.id ?? '';
    let changes = 0;
    d.on('document:change', () => {
      changes++;
    });
    d.tx((t) => void t.selectContentControlItem(id, 1));
    expect(changes).toBe(1);
    expect(
      d.bindings
        .list()
        .filter((b) => b.key === 'duplicate')
        .map((b) => [b.id, b.lastValue]),
    ).toEqual([
      [list[0]?.id, 'b'],
      [list[1]?.id, 'a'],
    ]);
    const again = await UltimateWord.load(await d.toDocx());
    expect(
      again.bindings
        .list()
        .filter((b) => b.key === 'duplicate')
        .map((b) => [b.value, b.lastValue]),
    ).toEqual([
      ['同名', 'b'],
      ['同名', 'a'],
    ]);
    d.undo();
    expect(
      d.bindings
        .list()
        .filter((b) => b.key === 'duplicate')
        .map((b) => b.lastValue),
    ).toEqual(['a', 'a']);
    d.redo();
    expect(d.bindings.list().find((b) => b.id === id)?.lastValue).toBe('b');
  });

  it('复选框拒绝自由编辑，全部替换跳过且不调用替换函数，勾选与导出回读仍一致', async () => {
    const d = await UltimateWord.load(template());
    const r = d.find('☐')[0] as DocRange;
    let events = 0;
    d.on('document:change', () => {
      events++;
    });
    expect(() => d.tx((t) => void t.replaceText(r, '☒'))).toThrow(/复选框只能/);
    const replace = vi.fn(() => '☒');
    expect(d.replaceAll('☐', replace)).toEqual({ replaced: 0, skipped: 1 });
    expect(replace).not.toHaveBeenCalled();
    expect(events).toBe(0);
    expect(d.canUndo).toBe(false);
    expect(d.bindings.list().find((b) => b.key === 'agree')).toMatchObject({
      checked: false,
      value: '☐',
      locked: false,
    });
    d.bindings.set('agree', true).apply();
    expect(events).toBe(1);
    const again = await UltimateWord.load(await d.toDocx());
    expect(again.bindings.list().find((b) => b.key === 'agree')).toMatchObject({
      checked: true,
      value: '☒',
    });
    expect(() => again.tx((t) => void t.deleteRange(again.find('☒')[0] as DocRange))).toThrow(/复选框只能/);
    expect(again.replaceAll('☒', '☐')).toEqual({ replaced: 0, skipped: 1 });
    d.undo();
    expect(d.find('☐')).toHaveLength(1);
    d.redo();
    expect(d.find('☒')).toHaveLength(1);
  });

  it('下拉框拒绝直接文字事务且不派发事件，填写与导出回读后仍受保护', async () => {
    const d = await UltimateWord.load(template());
    let events = 0;
    d.on('document:change', () => {
      events++;
    });
    expect(() => d.tx((t) => void t.replaceText(d.find('一般')[0] as DocRange, '紧急'))).toThrow(
      /下拉框只能/,
    );
    expect(events).toBe(0);
    expect(d.canUndo).toBe(false);
    d.bindings.set('level', '1').apply();
    expect(events).toBe(1);
    expect(d.bindings.list().find((b) => b.key === 'level')).toMatchObject({
      value: '紧急',
      lastValue: '1',
      locked: false,
    });
    const again = await UltimateWord.load(await d.toDocx());
    expect(() => again.tx((t) => void t.insertText((again.find('紧急')[0] as DocRange).end, '改'))).toThrow(
      /下拉框只能/,
    );
    expect(again.replaceAll('紧急', '一般')).toEqual({ replaced: 0, skipped: 1 });
    d.undo();
    expect(d.find('一般')).toHaveLength(1);
    d.redo();
    expect(d.find('紧急')).toHaveLength(1);
  });

  it('查找替换跳过下拉框并同步组合框保存值，事件只触发一次，撤销 / 导出使用同一状态', async () => {
    const d = await UltimateWord.load(
      template(
        `<w:p>${control('combo', '一般', '<w:comboBox w:lastValue="2"><w:listItem w:displayText="紧急" w:value="1"/><w:listItem w:displayText="一般" w:value="2"/></w:comboBox>')}</w:p>`,
      ),
    );
    let events = 0;
    const stop = d.on('document:change', () => {
      events++;
      expect(
        d.bindings
          .list()
          .filter((b) => b.items)
          .map((b) => b.lastValue),
      ).toEqual([undefined, '1']);
    });
    expect(d.replaceAll('一般', '紧急')).toEqual({ replaced: 1, skipped: 1 });
    expect(events).toBe(1);
    const again = await UltimateWord.load(await d.toDocx());
    expect(
      again.bindings
        .list()
        .filter((b) => b.items)
        .map((b) => [b.value, b.lastValue]),
    ).toEqual([
      ['一般', undefined],
      ['紧急', '1'],
    ]);
    stop.dispose();
    d.undo();
    expect(
      d.bindings
        .list()
        .filter((b) => b.items)
        .map((b) => b.lastValue),
    ).toEqual([undefined, '2']);
    d.redo();
    expect(
      d.bindings
        .list()
        .filter((b) => b.items)
        .map((b) => b.lastValue),
    ).toEqual([undefined, '1']);
  });

  it.each(['<w:text/>', '<w:richText/>'])(
    '文字控件 %s 普通替换立即清占位状态，事件、历史与导出一致',
    async (type) => {
      const d = await UltimateWord.load(
        template(
          `<w:p>${control('field', '占位文字', `${type}<w:showingPlcHdr/><w:dataBinding w:xpath="/field"/>`)}</w:p>`,
        ),
      );
      let changes = 0;
      const sub = d.on('document:change', () => {
        changes++;
        expect(d.bindings.list().find((b) => b.key === 'field')).toMatchObject({
          value: '张三',
          placeholder: false,
        });
      });
      expect(d.replaceAll('占位文字', '张三')).toEqual({ replaced: 1, skipped: 0 });
      expect(changes).toBe(1);
      sub.dispose();
      const current = d.bindings.list().find((b) => b.key === 'field');
      const again = await UltimateWord.load(await d.toDocx());
      expect(again.bindings.list().find((b) => b.key === 'field')).toMatchObject({
        value: '张三',
        placeholder: false,
      });
      d.undo();
      expect(d.bindings.list().find((b) => b.key === 'field')).toMatchObject({
        value: '占位文字',
        placeholder: true,
      });
      d.redo();
      expect(d.bindings.list().find((b) => b.key === 'field')).toEqual(current);
    },
  );

  it.each(['dropDownList', 'comboBox', 'date', 'checkbox', 'text', 'richText'])(
    '%s 显式填相同值解除绑定，事件只派发一次，重复提交无修改，历史与导出一致',
    async (type) => {
      const pr =
        type === 'date'
          ? '<w:date w:fullDate="2026-10-13T00:00:00Z"><w:dateFormat w:val="yyyy年M月d日"/></w:date>'
          : type === 'checkbox'
            ? '<w14:checkbox/>'
            : type === 'text' || type === 'richText'
              ? `<w:${type}/>`
              : `<w:${type} w:lastValue="2"><w:listItem w:displayText="一般" w:value="2"/></w:${type}>`;
      const text = type === 'date' ? '2026年10月13日' : type === 'checkbox' ? '☐' : '一般';
      const d = await UltimateWord.load(
        template(`<w:p>${control('bound', text, `${pr}<w:dataBinding w:xpath="/bound"/>`)}</w:p>`),
      );
      const parts = async () => unzip(new Uint8Array(await (await d.toDocx()).arrayBuffer()));
      const original = await parts();
      const xmlOf = async () => new TextDecoder().decode((await parts()).get('word/document.xml'));
      let changes = 0;
      let layouts = 0;
      d.on('document:change', () => changes++);
      d.on('layout:done', () => layouts++);
      const value =
        type === 'date'
          ? { date: '2026-10-13' }
          : type === 'checkbox'
            ? false
            : type === 'text' || type === 'richText'
              ? text
              : '2';
      expect(d.bindings.set('bound', value).apply()).toBeDefined();
      expect(await xmlOf()).not.toContain('dataBinding');
      expect(changes).toBe(1);
      expect(layouts).toBe(1);
      expect(d.bindings.pending.size).toBe(0);
      const current = d.bindings.list().find((b) => b.key === 'bound');
      expect(current?.value).toBe(text);
      expect(d.bindings.set('bound', value).apply()).toBeUndefined();
      expect(changes).toBe(1);
      expect(layouts).toBe(1);
      const again = await UltimateWord.load(await d.toDocx());
      expect(again.bindings.list().find((b) => b.key === 'bound')).toEqual(current);
      d.undo();
      const restored = await parts();
      for (const [part, content] of original) expect(restored.get(part)).toEqual(content);
      d.redo();
      expect(await xmlOf()).not.toContain('dataBinding');
      expect(d.bindings.list().find((b) => b.key === 'bound')).toEqual(current);
    },
  );

  it.each(['dropDownList', 'comboBox'])(
    '%s 填当前摘要不能绕过对象校验，失败保留待填值与历史',
    async (type) => {
      const list = `<w:${type} w:lastValue="2"><w:listItem w:displayText="一般" w:value="2"/></w:${type}>`;
      const d = await UltimateWord.load(
        template(
          `<w:p>${control('complex', '一般', `${list}<w:dataBinding w:xpath="/complex"/>`).replace('<w:t>一般</w:t>', '<w:t>一般</w:t><w:drawing/>')}</w:p>`,
        ),
      );
      d.bindings.set('applicant', '张三').apply();
      const original = new Uint8Array(await (await d.toDocx()).arrayBuffer());
      let events = 0;
      let layouts = 0;
      d.on('document:change', () => events++);
      d.on('layout:done', () => layouts++);
      d.bindings.set('applicant', '李四').set('complex', '2');
      expect(() => d.bindings.apply()).toThrow();
      expect(events).toBe(0);
      expect(layouts).toBe(0);
      expect(d.bindings.pending.get('complex')).toBe('2');
      expect(d.bindings.pending.get('applicant')).toBe('李四');
      expect(d.find('张三')).toHaveLength(2);
      expect(d.find('李四')).toHaveLength(0);
      expect(new Uint8Array(await (await d.toDocx()).arrayBuffer())).toEqual(original);
      d.undo();
      expect(d.find('张三')).toHaveLength(0);
      d.redo();
      expect(d.find('张三')).toHaveLength(2);
    },
  );

  it('普通文字事务同步组合框自填值并清旧日期，整批修改同撤销与导出', async () => {
    const d = await UltimateWord.load(
      template(
        `<w:p>${control('combo', '原值', '<w:comboBox w:lastValue="old"/>')}</w:p><w:p>${control('issued', '2026年10月13日', '<w:date w:fullDate="2026-10-13T00:00:00Z"><w:dateFormat w:val="yyyy年M月d日"/></w:date>')}</w:p>`,
      ),
    );
    d.tx((t) => {
      const combo = d.find('原值')[0];
      const date = d.find('2026年10月13日')[0];
      if (combo === undefined || date === undefined) throw new Error('样本缺少控件文字');
      t.insertText(combo.end, '待确认');
      t.insertText(date.end, '待确认');
    });
    expect(d.bindings.list().find((b) => b.key === 'combo')).toMatchObject({
      value: '原值待确认',
      lastValue: '原值待确认',
    });
    expect(d.bindings.list().find((b) => b.key === 'issued')?.fullDate).toBeUndefined();
    const again = await UltimateWord.load(await d.toDocx());
    expect(again.bindings.list().find((b) => b.key === 'combo')?.lastValue).toBe('原值待确认');
    expect(again.bindings.list().find((b) => b.key === 'issued')?.fullDate).toBeUndefined();
    d.undo();
    expect(d.bindings.list().find((b) => b.key === 'combo')?.lastValue).toBe('old');
    expect(d.bindings.list().find((b) => b.key === 'issued')?.fullDate).toBe('2026-10-13T00:00:00Z');
    d.redo();
    expect(d.bindings.list().find((b) => b.key === 'combo')?.lastValue).toBe('原值待确认');
    expect(d.bindings.list().find((b) => b.key === 'issued')?.fullDate).toBeUndefined();
  });

  it('日期对象按各自模板格式填同名控件，冻结 pending 值，混合填写同撤销且可导出再加载', async () => {
    const extra = `<w:p>${control('issued', '某日', '<w:date><w:dateFormat w:val="yyyy年M月d日"/><w:lid w:val="zh-CN"/></w:date>')}${control('issued', 'another', '<w:date><w:dateFormat w:val="dddd, MMMM d, yyyy"/><w:lid w:val="en-US"/></w:date>')}</w:p>`;
    const d = await UltimateWord.load(template(extra));
    const value = { date: '2026-10-13' };
    d.bindings.setMany({ applicant: '张三', agree: true, issued: value });
    value.date = '2020-01-01';
    expect(d.bindings.pending.get('issued')).toEqual({ date: '2026-10-13' });
    expect(Object.isFrozen(d.bindings.pending.get('issued'))).toBe(true);
    expect(d.bindings.apply()).toBeDefined();
    const dates = d.bindings.list().filter((b) => b.key === 'issued');
    expect(dates.map((b) => b.value)).toEqual(['2026年10月13日', 'Tuesday, October 13, 2026']);
    expect(dates.map((b) => b.fullDate)).toEqual(['2026-10-13T00:00:00Z', '2026-10-13T00:00:00Z']);
    d.undo();
    expect(
      d.bindings
        .list()
        .filter((b) => b.key === 'issued')
        .map((b) => b.value),
    ).toEqual(['某日', 'another']);
    expect(d.bindings.list().find((b) => b.key === 'applicant')?.value).toBe('单击输入');
    expect(d.bindings.list().find((b) => b.key === 'agree')?.checked).toBe(false);
    d.redo();
    const again = await UltimateWord.load(await d.toDocx());
    expect(
      again.bindings
        .list()
        .filter((b) => b.key === 'issued')
        .map((b) => [b.value, b.fullDate]),
    ).toEqual(dates.map((b) => [b.value, b.fullDate]));
    d.bindings.set('issued', { date: '2026-10-13' });
    expect(d.bindings.apply()).toBeUndefined();
  });

  it('日期局部加粗后仍可批量填写，显示、事件、历史与导出回读一致', async () => {
    const d = await UltimateWord.load(
      template(
        `<w:p>${control('issued', '2026年10月13日', '<w:date w:fullDate="2026-10-13T00:00:00Z"><w:dateFormat w:val="yyyy年M月d日"/></w:date>')}</w:p>`,
      ),
    );
    d.tx((t) => {
      t.setRunProps(d.find('2026年')[0] as DocRange, { bold: true });
    });
    const before = d.bindings.list().find((b) => b.key === 'issued');
    let changes = 0;
    const sub = d.on('document:change', () => {
      changes++;
    });
    d.bindings.set('issued', { date: '2024-02-29' }).apply();
    expect(changes).toBe(1);
    expect(d.bindings.pending.size).toBe(0);
    const after = d.bindings.list().find((b) => b.key === 'issued');
    expect(after).toMatchObject({ value: '2024年2月29日', fullDate: '2024-02-29T00:00:00Z' });
    expect(d.find('2024年2月29日')[0]).toEqual(after?.range);
    d.undo();
    expect(d.bindings.list().find((b) => b.key === 'issued')).toEqual(before);
    d.redo();
    expect(d.bindings.list().find((b) => b.key === 'issued')).toEqual(after);
    sub.dispose();
    const again = await UltimateWord.load(await d.toDocx());
    expect(again.bindings.list().find((b) => b.key === 'issued')).toMatchObject({
      value: '2024年2月29日',
      fullDate: '2024-02-29T00:00:00Z',
    });
  });

  it('日期对象验证全部同名控件；批量日期填写失败时回滚并保留 pending', async () => {
    const pr = '<w:date><w:dateFormat w:val="yyyy-MM-dd"/></w:date>';
    const d = await UltimateWord.load(template(`<w:p>${control('issued', '某日', pr)}</w:p>`));
    expect(() => d.bindings.set('issued', { date: '2026-02-29' })).toThrow(/日期/);
    expect(() => d.bindings.set('applicant', { date: '2026-10-13' })).toThrow(/字符串/);
    expect(d.bindings.pending.size).toBe(0);
    const mixed = await UltimateWord.load(
      template(`<w:p>${control('issued', '某日', pr)}${control('issued', '文字', '<w:text/>')}</w:p>`),
    );
    expect(() => mixed.bindings.set('issued', { date: '2026-10-13' })).toThrow(/字符串/);
    expect(mixed.bindings.pending.size).toBe(0);
    d.bindings.setMany({ issued: { date: '2026-10-13' }, level: '没有的选项' });
    expect(() => d.bindings.apply()).toThrow();
    expect(d.bindings.list().find((b) => b.key === 'issued')?.value).toBe('某日');
    expect(d.bindings.list().find((b) => b.key === 'issued')?.fullDate).toBeUndefined();
    expect(d.bindings.pending.get('issued')).toEqual({ date: '2026-10-13' });
    expect(d.canUndo).toBe(false);
  });

  it('名字拼错、复选框在 set 时就抛；下拉框没有的选项在 apply 时整批回滚、值留着', async () => {
    const d = await UltimateWord.load(template());
    expect(() => d.bindings.set('aplicant', '张三')).toThrow(/没有坑位/);
    expect(() => d.bindings.set('agree', '☑')).toThrow(/checkbox/);
    d.bindings.set('applicant', '张三').set('level', '特急').set('agree', true);
    expect(() => d.bindings.apply()).toThrow();
    expect(d.find('张三')).toHaveLength(0);
    expect(d.canUndo).toBe(false);
    expect(d.bindings.pending.get('applicant')).toBe('张三');
    expect(d.bindings.pending.get('agree')).toBe(true);
    expect(d.bindings.list().find((b) => b.key === 'agree')?.checked).toBe(false);
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

  it('复选框填布尔值，同名控件各用自己的符号；混合填写一次重排，撤销 / 重做与导出保持状态', async () => {
    const d = await UltimateWord.load(
      template(
        `<w:p>${control('agree', '□', '<w14:checkbox><w14:checkedState w14:val="2611" w14:font="宋体"/><w14:uncheckedState w14:val="25A1" w14:font="宋体"/></w14:checkbox>')}</w:p>`,
      ),
    );
    const states = (doc: UwDocument) =>
      doc.bindings
        .list()
        .filter((b) => b.key === 'agree')
        .map(({ checked, value }) => ({ checked, value }));
    expect(states(d)).toEqual([
      { checked: false, value: '☐' },
      { checked: false, value: '□' },
    ]);
    let changes = 0;
    let layouts = 0;
    d.on('document:change', () => {
      changes++;
    });
    d.on('layout:done', () => {
      layouts++;
    });
    expect(d.bindings.setMany({ applicant: '张三', agree: true, unused: false })).toEqual({
      skipped: ['unused'],
    });
    expect(d.bindings.pending.get('agree')).toBe(true);
    expect(d.bindings.apply()).toBeDefined();
    expect(changes).toBe(1);
    expect(layouts).toBe(1);
    expect(states(d)).toEqual([
      { checked: true, value: '☒' },
      { checked: true, value: '☑' },
    ]);
    expect(d.bindings.pending.size).toBe(0);
    expect(states(await UltimateWord.load(await d.toDocx()))).toEqual(states(d));
    d.bindings.set('agree', true);
    expect(d.bindings.apply()).toBeUndefined();
    expect(changes).toBe(1);
    expect(layouts).toBe(1);
    d.undo();
    expect(states(d)).toEqual([
      { checked: false, value: '☐' },
      { checked: false, value: '□' },
    ]);
    d.redo();
    expect(states(d)).toEqual([
      { checked: true, value: '☒' },
      { checked: true, value: '☑' },
    ]);
    d.bindings.set('agree', false).apply();
    expect(states(d)).toEqual([
      { checked: false, value: '☐' },
      { checked: false, value: '□' },
    ]);
  });

  it('复选框拒绝字符串，文字坑位拒绝布尔值；同名混合类型不会留下待提交值', async () => {
    const d = await UltimateWord.load(template());
    expect(() => d.bindings.set('agree', 'false')).toThrow(/布尔/);
    expect(() => d.bindings.set('applicant', true)).toThrow(/字符串/);
    expect(d.bindings.pending.size).toBe(0);
    const mixed = await UltimateWord.load(template(`<w:p>${control('agree', '文字')}</w:p>`));
    expect(() => mixed.bindings.set('agree', true)).toThrow(/字符串/);
    expect(mixed.bindings.pending.size).toBe(0);
  });

  it('列表报告继承的内容锁，复选框与文字坑位都不能绕过外层锁，失败后保留待提交值', async () => {
    const d = await UltimateWord.load(
      template(
        `<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:p>${control('lockedAgree', '☐', '<w14:checkbox/>')}${control('lockedName', '原名', '<w:text/>')}</w:p></w:sdtContent></w:sdt>`,
      ),
    );
    const locked = d.bindings.list().filter((b) => b.key.startsWith('locked'));
    expect(locked.map((b) => [b.key, b.locked])).toEqual([
      ['lockedAgree', true],
      ['lockedName', true],
    ]);
    d.bindings.set('lockedAgree', true);
    expect(() => d.bindings.apply()).toThrow(/锁定/);
    expect(d.bindings.pending.get('lockedAgree')).toBe(true);
    expect(d.bindings.list().find((b) => b.key === 'lockedAgree')?.checked).toBe(false);
    d.bindings.discard();
    d.bindings.set('lockedName', '张三');
    expect(() => d.bindings.apply()).toThrow(/锁定/);
    expect(d.bindings.list().find((b) => b.key === 'lockedName')?.value).toBe('原名');
    expect(d.canUndo).toBe(false);
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
  function revisedDocx(bodyXml?: string): Uint8Array {
    const enc = new TextEncoder();
    const ct = (part: string, type: string) =>
      `<Override PartName="/word/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`;
    const rel = (id: string, type: string, target: string) =>
      `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
    const body =
      bodyXml ??
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

  it('段落标记被删：最终状态排成一段、能在并进来的字里打字，所有标记下是两段；被删的行最终状态不占位', async () => {
    const d = await UltimateWord.load(
      revisedDocx(
        '<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="张三"/></w:rPr></w:pPr><w:r><w:t>前半</w:t></w:r></w:p>' +
          '<w:p><w:r><w:t>后半</w:t></w:r></w:p>' +
          '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
          '<w:tr><w:tc><w:p><w:r><w:t>留行</w:t></w:r></w:p></w:tc></w:tr>' +
          '<w:tr><w:trPr><w:del w:id="2" w:author="张三"/></w:trPr><w:tc><w:p><w:del w:id="3" w:author="张三"><w:r><w:delText>删行</w:delText></w:r></w:del></w:p></w:tc></w:tr>' +
          '</w:tbl><w:p/>',
      ),
    );
    const shape = () =>
      d.layout.pages[0]?.blocks.map((b) =>
        b.kind === 'paragraph' ? b.lines.length : `rows:${b.rows.length}`,
      );
    expect(shape()).toEqual([1, 'rows:1', 1]);
    expect(painted(d).body).toBe('前半后半');
    const [hit] = d.find('前半');
    if (hit === undefined) throw new Error('找不到「前半」');
    d.tx((t) => void t.insertText(hit.end, '甲'));
    expect(painted(d).body).toBe('前半甲后半');
    d.setRevisionView('markup');
    expect(shape()).toEqual([1, 1, 'rows:2', 1]);
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

  it('纯结构修订也列出：被删标记与整行不依赖字上的 w:del，三种视图与导出再加载的记录相同', async () => {
    const d = await UltimateWord.load(
      revisedDocx(
        '<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="标记作者"/></w:rPr></w:pPr><w:r><w:t>前段</w:t></w:r></w:p>' +
          '<w:p><w:r><w:t>后段</w:t></w:r></w:p>' +
          '<w:tbl><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>' +
          '<w:tr><w:trPr><w:del w:id="2" w:author="行作者"/></w:trPr>' +
          '<w:tc><w:p><w:r><w:t>甲格</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/>',
      ),
    );
    const summary = (doc: UwDocument) =>
      doc.revisions().map(({ target, kind, author, text }) => ({ target, kind, author, text }));
    const expected = [
      { target: 'paragraph-mark', kind: 'delete', author: '标记作者', text: '\n' },
      { target: 'row', kind: 'delete', author: '行作者', text: '甲格\t' },
    ];
    const list = d.revisions();
    expect(summary(d)).toEqual(expected);
    for (const view of ['markup', 'original', 'final'] as const) {
      d.setRevisionView(view);
      expect(d.revisions()).toEqual(list);
    }
    const row = list[1];
    if (row === undefined) throw new Error('缺少整行修订');
    expect(row.range).toEqual(d.rangeOf(row.rowId ?? ''));
    expect(d.compare(row.range.start, row.range.end)).toBe(-1);
    expect(summary(await UltimateWord.load(await d.toDocx()))).toEqual(expected);
  });

  it('段落标记的定位随拆段移动、随撤销恢复；编辑后导出不会复制出第二条修订', async () => {
    const d = await UltimateWord.load(
      revisedDocx(
        '<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="张三"/></w:rPr></w:pPr><w:r><w:t>甲乙</w:t></w:r></w:p>',
      ),
    );
    const before = d.revisions();
    const [hit] = d.find('甲');
    if (hit === undefined) throw new Error('缺少拆段位置');
    d.tx((t) => void t.splitParagraph(hit.end));
    const after = d.revisions();
    expect(after).toHaveLength(1);
    expect(after[0]?.target).toBe('paragraph-mark');
    expect(after[0]?.paragraphId).not.toBe(before[0]?.paragraphId);
    const end = d.rangeOf(after[0]?.paragraphId ?? '')?.end;
    expect(after[0]?.range).toEqual({ start: end, end });
    const loaded = await UltimateWord.load(await d.toDocx());
    expect(loaded.revisions()).toHaveLength(1);
    expect(loaded.revisions()[0]).toMatchObject({ target: 'paragraph-mark', kind: 'insert', author: '张三' });
    d.undo();
    expect(d.revisions()).toEqual(before);
    d.redo();
    expect(d.revisions()).toEqual(after);
  });

  it('替换跳过首字落在被删除修订里的命中，不让整批回滚', async () => {
    const d = await UltimateWord.load(revisedDocx(), { revisions: 'markup' });
    expect(d.replaceAll(/删掉|结尾/, '改')).toEqual({ replaced: 1, skipped: 1 });
    expect(painted(d).body).toBe('原文删掉新增改');
  });
});
