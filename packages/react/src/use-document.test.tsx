import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import type { LoadSource } from 'ultimate-word';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import type { DocumentState } from './use-document.ts';
import { useDocument } from './use-document.ts';

// jsdom 环境下 import.meta.url 是 http:// 的，拼不出 file 路径；vitest 的 cwd 是包目录
const bytes = new Uint8Array(readFileSync(resolve('../../apps/fidelity/fixtures/gongwen-01.docx')));

beforeAll(() => {
  (globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
});

let root: Root | undefined;
let last: DocumentState | undefined;
function Probe({ source }: { source: LoadSource | null }) {
  last = useDocument(source);
  return null;
}
function render(source: LoadSource | null) {
  root ??= createRoot(document.body.appendChild(document.createElement('div')));
  act(() => root?.render(<Probe source={source} />));
}
const settle = () => act(async () => new Promise((r) => setTimeout(r, 50)));
afterEach(() => {
  act(() => root?.unmount());
  root = undefined;
});

describe('useDocument', () => {
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
