/**
 * 组件的判据是**接线**：什么时候 mount / dispose、zoom 走不走 setZoom、overlays 的 key 调和
 * 通不通到 `view.overlay()`。视图本身用假的（记下每次调用）—— 字摆在哪由 ultimate-word 的
 * 真值断言负责，这里挂真视图只会把 jsdom 没有的 getScreenCTM 之类拖进来。
 */
import { act, createRef, useState } from 'react';
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
      const root = document.createElement('div');
      target.replaceChildren(root);
      const view: FakeView = {
        root,
        zoom: 1,
        options,
        overlays: [],
        decorations: [],
        setZoom: vi.fn<UwView['setZoom']>(),
        dispose: vi.fn(() => root.remove()),
        locate: () => null,
        rectsOf: () => [],
        caretRect: () => null,
        scrollTo: () => false,
        print: () => undefined,
        on: () => ({ dispose: () => undefined }),
        decorate: (range) => {
          const item = { range, dispose: vi.fn() };
          view.decorations.push(item);
          return { update: vi.fn(), dispose: item.dispose };
        },
        overlay: (_position, element) => {
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
});

describe('useDecoration', () => {
  function Highlight({ view, range, cls }: { view: UwView | null; range: DocRange | null; cls: string }) {
    useDecoration(view, range, { className: cls });
    return null;
  }
  it('视图与 range 都在时挂一条装饰；按值比，同样的 range 新对象不重挂；变了才换', () => {
    const { doc, views } = fakeDocument();
    render(<UltimateWordView doc={doc} />);
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
});
