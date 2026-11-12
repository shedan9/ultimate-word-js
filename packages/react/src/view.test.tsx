/**
 * 组件的判据是**接线**：什么时候 mount / dispose、zoom 走不走 setZoom、overlays 的 key 调和
 * 通不通到 `view.overlay()`。视图本身用假的（记下每次调用）—— 字摆在哪由 ultimate-word 的
 * 真值断言负责，这里挂真视图只会把 jsdom 没有的 getScreenCTM 之类拖进来。
 */
import { act, createRef, StrictMode, useState } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import type { DocRange, OverlayHandle, UwDocument, UwView, ViewOptions } from 'ultimate-word';
import { afterEach, beforeAll, describe, expect, it, type Mock, vi } from 'vitest';
import { useDecoration } from './use-decoration.ts';
import type { ReactOverlaySpec } from './view.tsx';
import { UltimateWordView } from './view.tsx';

beforeAll(() => {
  (globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
});

interface FakeView extends UwView {
  options: ViewOptions;
  setZoom: Mock<UwView['setZoom']>;
  dispose: Mock<UwView['dispose']>;
  overlays: { element: HTMLElement; handle: OverlayHandle; update: Mock<OverlayHandle['update']> }[];
  decorations: { range: DocRange; dispose: Mock<() => void> }[];
}

function fakeDocument() {
  const views: FakeView[] = [];
  const doc = {
    mount: vi.fn((target: Element, options: ViewOptions = {}) => {
      let disposed = false;
      const root = document.createElement('div');
      target.replaceChildren(root);
      const view: FakeView = {
        root,
        zoom: 1,
        options,
        overlays: [],
        decorations: [],
        setZoom: vi.fn<UwView['setZoom']>(),
        dispose: vi.fn(() => {
          disposed = true;
          root.remove();
        }),
        locate: () => null,
        rectsOf: () => [],
        caretRect: () => null,
        scrollTo: () => false,
        print: () => undefined,
        toPNG: () => Promise.reject(new Error('假视图')),
        on: () => ({ dispose: () => undefined }),
        decorate: (range) => {
          if (disposed) throw new Error('视图已销毁');
          const item = { range, dispose: vi.fn() };
          view.decorations.push(item);
          return { update: vi.fn(), dispose: item.dispose };
        },
        overlay: (_position, element) => {
          if (disposed) throw new Error('视图已销毁');
          if (view.overlays.some((o) => o.element === element)) throw new Error('重复挂同一个元素');
          root.append(element);
          const update = vi.fn<OverlayHandle['update']>();
          const handle: OverlayHandle = {
            update,
            dispose: vi.fn(() => {
              view.overlays = view.overlays.filter((o) => o.element !== element);
              element.remove();
            }),
          };
          view.overlays.push({ element, handle, update });
          return handle;
        },
      };
      views.push(view);
      return view;
    }),
  };
  return { doc: doc as unknown as UwDocument, views };
}

const at = (nodeId: string, offset = 0) => ({ nodeId, contentIndex: 0, offset });
let root: Root | undefined;
let container: HTMLDivElement | undefined;
function render(node: React.ReactNode) {
  container ??= document.body.appendChild(document.createElement('div'));
  root ??= createRoot(container);
  act(() => root?.render(node));
}
afterEach(() => {
  act(() => root?.unmount());
  root = undefined;
  container?.remove();
  container = undefined;
});

describe('<UltimateWordView>', () => {
  it.each(['preview', 'edit'] as const)('完整传递构造选项，包含模式 %s', (mode) => {
    const { doc, views } = fakeDocument();
    const options: ViewOptions = {
      mode,
      zoom: 'fit-width',
      pageGap: 8,
      textLayer: true,
      virtualize: false,
      overscan: 0,
      classPrefix: 'react-view',
      fontFamily: (family) => family,
      debug: false,
      printMode: 'inline',
    };
    render(<UltimateWordView doc={doc} {...options} />);
    expect(doc.mount).toHaveBeenCalledTimes(1);
    expect(views[0]?.options).toEqual(options);
  });

  it('预览与编辑模式切换会清理旧视图，ref 与通知更新，移除 mode 恢复默认模式', () => {
    const { doc, views } = fakeDocument();
    const ref = createRef<UwView | null>();
    const onViewChange = vi.fn();
    render(<UltimateWordView doc={doc} ref={ref} onViewChange={onViewChange} zoom={1.5} />);
    const notifications: (UwView | null | undefined)[] = [views[0]];
    for (const mode of ['edit', 'preview', 'edit', undefined] as const) {
      const previous = views.at(-1);
      const count = views.length;
      const options = mode === undefined ? {} : { mode };
      render(<UltimateWordView doc={doc} ref={ref} onViewChange={onViewChange} zoom={1.5} {...options} />);
      expect(views).toHaveLength(count + 1);
      const current = views.at(-1);
      expect(previous?.dispose).toHaveBeenCalledTimes(1);
      expect(current?.options).toEqual({ zoom: 1.5, ...options });
      expect(current?.setZoom).not.toHaveBeenCalled();
      expect(ref.current).toBe(current);
      notifications.push(null, current);
      expect(onViewChange.mock.calls.map(([value]) => value)).toEqual(notifications);
    }
    act(() => root?.unmount());
    expect(views.at(-1)?.dispose).toHaveBeenCalledTimes(1);
    expect(ref.current).toBeNull();
    expect(onViewChange).toHaveBeenLastCalledWith(null);
  });

  it('编辑模式未变时重渲染不重挂，缩放仍只更新当前视图', () => {
    const { doc, views } = fakeDocument();
    const ref = createRef<UwView | null>();
    render(<UltimateWordView doc={doc} mode="edit" ref={ref} />);
    const first = ref.current;
    render(<UltimateWordView doc={doc} mode="edit" ref={ref} className="updated" />);
    render(<UltimateWordView doc={doc} mode="edit" ref={ref} zoom={1.5} />);
    expect(doc.mount).toHaveBeenCalledTimes(1);
    expect(ref.current).toBe(first);
    expect(views[0]?.dispose).not.toHaveBeenCalled();
    expect(views[0]?.setZoom).toHaveBeenCalledExactlyOnceWith(1.5);
  });

  it('挂上就 mount 进容器，卸下就 dispose；ref 与 onViewChange 拿到的是同一个视图', () => {
    const { doc, views } = fakeDocument();
    const ref = createRef<UwView | null>();
    const onViewChange = vi.fn();
    render(<UltimateWordView doc={doc} ref={ref} onViewChange={onViewChange} zoom="fit-width" pageGap={8} />);
    expect(doc.mount).toHaveBeenCalledTimes(1);
    expect(views[0]?.options).toEqual({ zoom: 'fit-width', pageGap: 8 });
    expect(ref.current).toBe(views[0]);
    expect(onViewChange).toHaveBeenLastCalledWith(views[0]);
    expect(container?.firstElementChild?.firstElementChild).toBe(views[0]?.root);
    act(() => root?.unmount());
    expect(views[0]?.dispose).toHaveBeenCalledTimes(1);
    expect(onViewChange).toHaveBeenLastCalledWith(null);
  });

  it('zoom 变了只 setZoom，不重挂；构造选项变了才 dispose 再 mount', () => {
    const { doc, views } = fakeDocument();
    render(<UltimateWordView doc={doc} zoom={1} />);
    render(<UltimateWordView doc={doc} zoom={1.5} />);
    expect(doc.mount).toHaveBeenCalledTimes(1);
    expect(views[0]?.setZoom).toHaveBeenCalledWith(1.5);
    render(<UltimateWordView doc={doc} zoom={1.5} textLayer={false} />);
    expect(views[0]?.dispose).toHaveBeenCalledTimes(1);
    expect(doc.mount).toHaveBeenCalledTimes(2);
    // 重挂那一趟带上了当时的 zoom，不必再 setZoom 一次
    expect(views[1]?.options).toEqual({ zoom: 1.5, textLayer: false });
    expect(views[1]?.setZoom).not.toHaveBeenCalled();
  });

  it('换 doc 就重挂到新文档上', () => {
    const a = fakeDocument();
    const b = fakeDocument();
    render(<UltimateWordView doc={a.doc} />);
    render(<UltimateWordView doc={b.doc} />);
    expect(a.views[0]?.dispose).toHaveBeenCalledTimes(1);
    expect(b.doc.mount).toHaveBeenCalledTimes(1);
  });

  it('overlays：内容是真正的 React 子树，按 key 调和，锚点变了 update，重挂后气泡里的状态不丢', () => {
    const { doc, views } = fakeDocument();
    function Bubble({ label }: { label: string }) {
      const [n, setN] = useState(0);
      return (
        <button type="button" data-label={label} onClick={() => setN(n + 1)}>
          {label}:{n}
        </button>
      );
    }
    const specs = (anchorOffset: number, keys: string[]): ReactOverlaySpec[] =>
      keys.map((key) => ({ key, anchor: at('r0', anchorOffset), render: () => <Bubble label={key} /> }));
    render(<UltimateWordView doc={doc} overlays={specs(0, ['a', 'b'])} />);
    const view = views[0] as FakeView;
    expect(view.overlays).toHaveLength(2);
    const button = view.root.querySelector('[data-label="a"]') as HTMLButtonElement;
    expect(button.isConnected).toBe(true);
    act(() => button.click());
    expect(button.textContent).toBe('a:1');
    // 锚点变了 → update()，元素与 state 不动
    render(<UltimateWordView doc={doc} overlays={specs(5, ['a', 'b'])} />);
    expect(view.overlays[0]?.update).toHaveBeenCalledWith(at('r0', 5));
    expect(view.root.querySelector('[data-label="a"]')).toBe(button);
    expect(button.textContent).toBe('a:1');
    // 摘掉 b
    render(<UltimateWordView doc={doc} overlays={specs(5, ['a'])} />);
    expect(view.overlays).toHaveLength(1);
    expect(view.root.querySelector('[data-label="b"]')).toBeNull();
    // 构造选项变了 → 视图重挂，overlay 在新视图上重新挂，气泡里的 state 还在
    render(<UltimateWordView doc={doc} overlays={specs(5, ['a'])} pageGap={0} />);
    const next = views[1] as FakeView;
    expect(next.overlays).toHaveLength(1);
    expect(next.root.querySelector('[data-label="a"]')).toBe(button);
    expect(button.textContent).toBe('a:1');
    // 模式也是构造选项，切换时只换底层视图，不应卸载宿主中的 React 子树。
    for (const mode of ['edit', 'preview'] as const) {
      render(<UltimateWordView doc={doc} overlays={specs(5, ['a'])} pageGap={0} mode={mode} />);
      const current = views.at(-1) as FakeView;
      expect(current.options.mode).toBe(mode);
      expect(current.overlays).toHaveLength(1);
      expect(current.root.querySelector('[data-label="a"]')).toBe(button);
      expect(button.textContent).toBe('a:1');
    }
  });

  it('overlay 的 key 重复直接抛', () => {
    const { doc } = fakeDocument();
    const spy = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    expect(() =>
      render(
        <UltimateWordView
          doc={doc}
          overlays={[
            { key: 'x', anchor: at('r0'), render: () => null },
            { key: 'x', anchor: at('r1'), render: () => null },
          ]}
        />,
      ),
    ).toThrow(/key 重复/);
    spy.mockRestore();
  });

  it('数字与同文字符串 key 各自保留气泡状态，换序、移除与重挂不混用身份', () => {
    const { doc, views } = fakeDocument();
    const errors = vi.spyOn(console, 'error');
    function Draft({ label }: { label: string }) {
      const [text, setText] = useState('草稿');
      return (
        <button type="button" data-label={label} onClick={() => setText(`${label}已修改`)}>
          {text}
        </button>
      );
    }
    const numeric: ReactOverlaySpec = { key: 1, anchor: at('r0'), render: () => <Draft label="数字" /> };
    const textual: ReactOverlaySpec = { key: '1', anchor: at('r1'), render: () => <Draft label="字符串" /> };
    render(<UltimateWordView doc={doc} overlays={[numeric, textual]} />);
    const first = views[0] as FakeView;
    const numberButton = first.root.querySelector('[data-label="数字"]') as HTMLButtonElement;
    const stringButton = first.root.querySelector('[data-label="字符串"]') as HTMLButtonElement;
    act(() => {
      numberButton.click();
      stringButton.click();
    });
    render(<UltimateWordView doc={doc} overlays={[textual, numeric]} />);
    expect(first.root.querySelector('[data-label="数字"]')).toBe(numberButton);
    expect(first.root.querySelector('[data-label="字符串"]')).toBe(stringButton);
    expect(numberButton.textContent).toBe('数字已修改');
    expect(stringButton.textContent).toBe('字符串已修改');
    render(<UltimateWordView doc={doc} overlays={[textual]} pageGap={8} />);
    const next = views[1] as FakeView;
    expect(next.overlays).toHaveLength(1);
    expect(numberButton.isConnected).toBe(false);
    expect(next.root.querySelector('button')).toBe(stringButton);
    expect(stringButton.textContent).toBe('字符串已修改');
    render(<UltimateWordView doc={doc} overlays={[numeric, textual]} pageGap={8} />);
    expect(next.root.querySelector('[data-label="数字"]')?.textContent).toBe('草稿');
    expect(next.root.querySelector('[data-label="字符串"]')).toBe(stringButton);
    expect(errors).not.toHaveBeenCalled();
    errors.mockRestore();
  });

  it.each(['doc', 'mode', 'pageGap'] as const)(
    '同次更新 %s 与 overlays 时只向新视图挂气泡，保留同 key 的草稿',
    (change) => {
      const first = fakeDocument();
      const second = fakeDocument();
      const ref = createRef<UwView | null>();
      function Draft() {
        const [text, setText] = useState('草稿');
        return (
          <button type="button" onClick={() => setText('已修改草稿')}>
            {text}
          </button>
        );
      }
      const kept: ReactOverlaySpec = { key: 'kept', anchor: at('r0'), render: () => <Draft /> };
      render(
        <UltimateWordView
          doc={first.doc}
          ref={ref}
          overlays={[kept, { key: 'removed', anchor: at('r1'), render: () => <span>待移除</span> }]}
        />,
      );
      const previous = first.views[0] as FakeView;
      const draft = previous.root.querySelector('button') as HTMLButtonElement;
      act(() => draft.click());
      const oldHandles = previous.overlays.map((item) => item.handle);
      const doc = change === 'doc' ? second.doc : first.doc;
      const options: ViewOptions =
        change === 'mode' ? { mode: 'edit' } : change === 'pageGap' ? { pageGap: 8 } : {};
      render(
        <UltimateWordView
          doc={doc}
          ref={ref}
          {...options}
          overlays={[
            { ...kept, anchor: at('r0', 3), placement: 'below', offset: { x: 0, y: 4 } },
            { key: 'added', anchor: at('r2'), render: () => <span>新批注</span> },
          ]}
        />,
      );
      const current = (change === 'doc' ? second.views[0] : first.views[1]) as FakeView;
      expect(previous.dispose).toHaveBeenCalledTimes(1);
      expect(oldHandles.every((handle) => vi.mocked(handle.dispose).mock.calls.length === 1)).toBe(true);
      expect(current.overlays).toHaveLength(2);
      expect(current.root.querySelector('button')).toBe(draft);
      expect(draft.textContent).toBe('已修改草稿');
      expect(current.root.textContent).toBe('已修改草稿新批注');
      expect(ref.current).toBe(current);
    },
  );

  it('StrictMode 重放挂载后气泡只属于存活视图，卸载清理全部句柄', () => {
    const { doc, views } = fakeDocument();
    const ref = createRef<UwView | null>();
    render(
      <StrictMode>
        <UltimateWordView
          doc={doc}
          ref={ref}
          overlays={[{ key: 'a', anchor: at('r0'), render: () => <span>批注</span> }]}
        />
      </StrictMode>,
    );
    expect(views).toHaveLength(2);
    const current = views[1] as FakeView;
    expect(views[0]?.dispose).toHaveBeenCalledTimes(1);
    expect(current.dispose).not.toHaveBeenCalled();
    expect(current.overlays).toHaveLength(1);
    expect(current.root.textContent).toBe('批注');
    expect(ref.current).toBe(current);
    const handle = current.overlays[0]?.handle;
    act(() => root?.unmount());
    expect(current.dispose).toHaveBeenCalledTimes(1);
    expect(handle?.dispose).toHaveBeenCalledTimes(1);
    expect(ref.current).toBeNull();
  });
});

describe('useDecoration', () => {
  function Highlight({ view, range, cls }: { view: UwView | null; range: DocRange | null; cls: string }) {
    useDecoration(view, range, { className: cls });
    return null;
  }
  it('视图与 range 都在时挂一条装饰；按值比，同样的 range 新对象不重挂；变了才换', () => {
    const { doc, views } = fakeDocument();
    doc.mount(document.createElement('div'));
    const view = views[0] as FakeView;
    const range = () => ({ start: at('r0'), end: at('r0', 2) });
    render(<Highlight view={view} range={range()} cls="hit" />);
    expect(view.decorations).toHaveLength(1);
    render(<Highlight view={view} range={range()} cls="hit" />);
    expect(view.decorations).toHaveLength(1);
    expect(view.decorations[0]?.dispose).not.toHaveBeenCalled();
    render(<Highlight view={view} range={{ start: at('r0'), end: at('r0', 3) }} cls="hit" />);
    expect(view.decorations[0]?.dispose).toHaveBeenCalledTimes(1);
    expect(view.decorations).toHaveLength(2);
    render(<Highlight view={view} range={null} cls="hit" />);
    expect(view.decorations[1]?.dispose).toHaveBeenCalledTimes(1);
  });

  it.each(['doc', 'mode', 'pageGap', 'hide'] as const)(
    '同次更新 %s 与高亮范围时跳过已销毁视图，向新视图挂装饰',
    (change) => {
      const first = fakeDocument();
      const second = fakeDocument();
      function Probe({ updated }: { updated: boolean }) {
        const [view, setView] = useState<UwView | null>(null);
        useDecoration(
          view,
          { start: at('r0'), end: at('r0', updated ? 3 : 2) },
          {
            className: updated ? 'new-hit' : 'hit',
          },
        );
        if (updated && change === 'hide') return null;
        return (
          <UltimateWordView
            doc={updated && change === 'doc' ? second.doc : first.doc}
            onViewChange={setView}
            {...(updated && change === 'mode' ? { mode: 'edit' } : {})}
            {...(updated && change === 'pageGap' ? { pageGap: 8 } : {})}
          />
        );
      }
      render(<Probe updated={false} />);
      const previous = first.views[0] as FakeView;
      expect(previous.decorations).toHaveLength(1);
      render(<Probe updated />);
      expect(previous.dispose).toHaveBeenCalledTimes(1);
      expect(previous.decorations).toHaveLength(1);
      expect(previous.decorations[0]?.dispose).toHaveBeenCalledTimes(1);
      if (change !== 'hide') {
        const current = (change === 'doc' ? second.views[0] : first.views[1]) as FakeView;
        expect(current.decorations).toHaveLength(1);
        expect(current.decorations[0]?.range.end.offset).toBe(3);
        act(() => root?.unmount());
        expect(current.decorations[0]?.dispose).toHaveBeenCalledTimes(1);
      }
    },
  );
});
