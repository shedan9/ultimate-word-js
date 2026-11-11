import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import type { LoadOptions, LoadSource, UwDocument } from 'ultimate-word';
import { UltimateWord } from 'ultimate-word';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type { DocumentState } from './use-document.ts';
import { useDocument } from './use-document.ts';

// jsdom 环境下 import.meta.url 是 http:// 的，拼不出 file 路径；vitest 的 cwd 是包目录
const bytes = new Uint8Array(readFileSync(resolve('../../apps/fidelity/fixtures/gongwen-01.docx')));

beforeAll(() => {
  (globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
});

let root: Root | undefined;
let last: DocumentState | undefined;
function Probe({ source, options }: { source: LoadSource | null; options: LoadOptions }) {
  last = useDocument(source, options);
  return null;
}
function render(source: LoadSource | null, options: LoadOptions = {}) {
  root ??= createRoot(document.body.appendChild(document.createElement('div')));
  act(() => root?.render(<Probe source={source} options={options} />));
}
const settle = () => act(async () => new Promise((r) => setTimeout(r, 50)));
afterEach(() => {
  act(() => root?.unmount());
  root = undefined;
  last = undefined;
  vi.restoreAllMocks();
});

describe('useDocument', () => {
  function addComment(doc = last?.doc) {
    if (!doc) throw new Error('样本没有加载完成');
    const hit = doc.find('通知')[0];
    if (!hit) throw new Error('样本缺少通知');
    const id = doc.addComment(hit, '批注回归');
    return doc.comments().find((comment) => comment.id === id);
  }

  it.each(['original', 'markup'] as const)('加载时传递修订视图 %s', async (revisions) => {
    render(bytes, { revisions });
    await settle();
    expect(last?.error).toBeNull();
    expect(last?.doc?.revisionView).toBe(revisions);
  });

  it('加载时传递默认作者，批注署名与导出回读一致', async () => {
    render(bytes, { author: '张三' });
    await settle();
    expect(addComment()).toMatchObject({ author: '张三', initials: '张' });
    const blob = await last?.doc?.toDocx();
    if (!blob) throw new Error('样本没有导出内容');
    const again = await UltimateWord.load(blob);
    expect(again.comments()[0]).toMatchObject({ author: '张三', initials: '张', text: '批注回归' });
  });

  it('同来源更改或移除修订选项会重新加载，等待时保留上一份文档', async () => {
    const load = vi.spyOn(UltimateWord, 'load');
    render(bytes, { revisions: 'original', author: '张三' });
    await settle();
    const first = last?.doc;
    render(bytes, { revisions: 'markup', author: '张三' });
    expect(last).toMatchObject({ doc: first, loading: true, error: null });
    expect(load.mock.calls[0]?.[1]?.signal?.aborted).toBe(true);
    await settle();
    expect(last?.doc).not.toBe(first);
    expect(last?.doc?.revisionView).toBe('markup');
    expect(addComment()?.author).toBe('张三');
    const second = last?.doc;
    render(bytes, { author: '张三' });
    await settle();
    expect(last?.doc).not.toBe(second);
    expect(last?.doc?.revisionView).toBe('final');
    expect(load).toHaveBeenCalledTimes(3);
  });

  it('同来源更改或移除默认作者会重新加载，不沿用旧署名', async () => {
    render(bytes, { revisions: 'original', author: '张三' });
    await settle();
    const first = last?.doc;
    render(bytes, { revisions: 'original', author: '李四' });
    expect(last).toMatchObject({ doc: first, loading: true });
    await settle();
    expect(last?.doc).not.toBe(first);
    expect(last?.doc?.revisionView).toBe('original');
    expect(addComment()).toMatchObject({ author: '李四', initials: '李' });
    render(bytes, { revisions: 'original' });
    await settle();
    expect(addComment()?.author).toBeUndefined();
    expect(last?.doc?.comments()).toHaveLength(1);
  });

  it('选项按字段比较，同值新对象不重新加载，字体与取消信号继续传递', async () => {
    const load = vi.spyOn(UltimateWord, 'load');
    const controller = new AbortController();
    const options: LoadOptions = {
      revisions: 'markup',
      author: '',
      fonts: UltimateWord.fonts.registry,
      signal: controller.signal,
    };
    render(bytes, options);
    await settle();
    const first = last?.doc;
    render(bytes, { ...options });
    await settle();
    expect(last?.doc).toBe(first);
    expect(load).toHaveBeenCalledTimes(1);
    const passed = load.mock.calls[0]?.[1];
    expect(passed).toMatchObject({ revisions: 'markup', author: '', fonts: options.fonts });
    expect(passed?.signal?.aborted).toBe(false);
    controller.abort();
    expect(passed?.signal?.aborted).toBe(true);
  });

  it.each(['resolve', 'reject'] as const)('选项变更取消旧加载，迟到的 %s 不覆盖新文档', async (outcome) => {
    const oldDoc = await UltimateWord.load(bytes, { revisions: 'original', author: '旧作者' });
    const newDoc = await UltimateWord.load(bytes, { revisions: 'markup', author: '新作者' });
    let resolveOld: ((doc: UwDocument) => void) | undefined;
    let rejectOld: ((error: Error) => void) | undefined;
    const pending = new Promise<UwDocument>((resolve, reject) => {
      resolveOld = resolve;
      rejectOld = reject;
    });
    const load = vi.spyOn(UltimateWord, 'load').mockReturnValueOnce(pending).mockResolvedValueOnce(newDoc);
    render(bytes, { revisions: 'original', author: '旧作者' });
    render(bytes, { revisions: 'markup', author: '新作者' });
    expect(load).toHaveBeenCalledTimes(2);
    expect(load.mock.calls[0]?.[1]?.signal?.aborted).toBe(true);
    await settle();
    expect(last).toEqual({ doc: newDoc, loading: false, error: null });
    await act(async () => {
      if (outcome === 'resolve') resolveOld?.(oldDoc);
      else rejectOld?.(new Error('旧选项加载失败'));
      await pending.catch(() => undefined);
    });
    expect(last).toEqual({ doc: newDoc, loading: false, error: null });
    expect(addComment()?.author).toBe('新作者');
    act(() => root?.unmount());
    root = undefined;
    expect(load.mock.calls[1]?.[1]?.signal?.aborted).toBe(true);
  });

  it('loading → doc；来源为 null 时空闲', async () => {
    render(bytes);
    expect(last).toMatchObject({ doc: null, loading: true, error: null });
    await settle();
    expect(last?.loading).toBe(false);
    expect(last?.doc?.pageCount).toBe(1);
    render(null);
    expect(last).toEqual({ doc: null, loading: false, error: null });
  });

  it('不是 zip 时 error 里是那个 UwError，不抛到渲染里', async () => {
    render(new Uint8Array([1, 2, 3]));
    await settle();
    expect(last?.doc).toBeNull();
    expect(last?.error).toBeInstanceOf(Error);
  });

  it('换来源时上一份文档先留着（loading 期间不闪白），回来后换成新的', async () => {
    render(bytes);
    await settle();
    const first = last?.doc;
    render(new Uint8Array(bytes));
    expect(last).toMatchObject({ doc: first, loading: true });
    await settle();
    expect(last?.doc).not.toBe(first);
    expect(last?.doc?.pageCount).toBe(1);
  });
});
