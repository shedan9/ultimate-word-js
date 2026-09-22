/**
 * 门面的判据不是真值残差 —— 字摆在哪由底下已标定完的几层决定，这里只验「接线没接错」：
 * 四种 `LoadSource` 都通到同一份布局、`find` 自动带上域求值结果、`rangeOf` 认 id 也认节点、
 * 随库度量包在模块加载时就已经在注册表里。跑在纯 Node 上，`mount()` 的那一半在
 * `apps/playground/tests/facade.html`。
 */
import { readFileSync } from 'node:fs';
import { FALLBACK_METRICS, FontRegistry } from '@uw/fonts';
import { beforeAll, describe, expect, it } from 'vitest';
import type { UwDocument } from './index.ts';
import { UltimateWord, UwError, UwErrorCode } from './index.ts';

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
