/**
 * 门面的判据不是真值残差 —— 字摆在哪由底下已标定完的几层决定，这里只验「接线没接错」：
 * 四种 `LoadSource` 都通到同一份布局、`find` 自动带上域求值结果、`rangeOf` 认 id 也认节点、
 * 随库度量包在模块加载时就已经在注册表里。跑在纯 Node 上，`mount()` 的那一半在
 * `apps/playground/tests/facade.html`。
 */
import { readFileSync } from 'node:fs';
import { FALLBACK_METRICS, FontRegistry } from '@uw/fonts';
import { unzip } from '@uw/ooxml';
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
