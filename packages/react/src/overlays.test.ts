import { describe, expect, it, vi } from 'vitest';
import type { OverlayEntry, OverlaySpec, ReconcileHost } from './overlays.ts';
import { disposeOverlays, reconcileOverlays } from './overlays.ts';

const at = (nodeId: string, offset = 0) => ({ nodeId, contentIndex: 0, offset });

function host(owner: object | null = {}) {
  const handles: {
    spec: OverlaySpec;
    update: ReturnType<typeof vi.fn>;
    dispose: ReturnType<typeof vi.fn>;
  }[] = [];
  const released: (string | number)[] = [];
  const h: ReconcileHost<OverlaySpec> = {
    owner,
    mount: vi.fn((spec: OverlaySpec) => {
      const handle = { spec, update: vi.fn(), dispose: vi.fn() };
      handles.push(handle);
      return handle;
    }),
    release: (key) => released.push(key),
  };
  return { h, handles, released };
}

describe('reconcileOverlays', () => {
  it('新 key 挂、消失的 key 摘并释放宿主、没变的一个都不动', () => {
    const current = new Map<string | number, OverlayEntry>();
    const { h, handles, released } = host();
    reconcileOverlays(
      current,
      [
        { key: 'a', anchor: at('r0') },
        { key: 'b', anchor: at('r1') },
      ],
      h,
    );
    expect(handles.map((x) => x.spec.key)).toEqual(['a', 'b']);
    reconcileOverlays(
      current,
      [
        { key: 'b', anchor: at('r1') },
        { key: 'c', anchor: at('r2') },
      ],
      h,
    );
    expect(handles[0]?.dispose).toHaveBeenCalledTimes(1);
    expect(released).toEqual(['a']);
    expect(handles[1]?.dispose).not.toHaveBeenCalled();
    expect(handles[1]?.update).not.toHaveBeenCalled();
    expect([...current.keys()]).toEqual(['b', 'c']);
  });

  it('锚点变了走 update()，placement / offset 变了只能摘了重挂', () => {
    const current = new Map<string | number, OverlayEntry>();
    const { h, handles, released } = host();
    reconcileOverlays(current, [{ key: 1, anchor: at('r0'), placement: 'above' }], h);
    reconcileOverlays(current, [{ key: 1, anchor: at('r0', 3), placement: 'above' }], h);
    expect(handles[0]?.update).toHaveBeenCalledWith(at('r0', 3));
    expect(handles).toHaveLength(1);
    reconcileOverlays(current, [{ key: 1, anchor: at('r0', 3), placement: 'below' }], h);
    expect(handles[0]?.dispose).toHaveBeenCalledTimes(1);
    expect(handles).toHaveLength(2);
    // 宿主元素还在用 —— key 没消失就不释放，portal 里的 React 子树因此不动
    expect(released).toEqual([]);
    reconcileOverlays(
      current,
      [{ key: 1, anchor: at('r0', 3), placement: 'below', offset: { x: 1, y: 2 } }],
      h,
    );
    expect(handles).toHaveLength(3);
    expect(h.mount).toHaveBeenLastCalledWith(expect.anything(), {
      placement: 'below',
      offset: { x: 1, y: 2 },
    });
  });

  it('视图换了（owner 不同）整组重挂；owner 为 null 时只摘不挂', () => {
    const current = new Map<string | number, OverlayEntry>();
    const first = host({});
    reconcileOverlays(current, [{ key: 'a', anchor: at('r0') }], first.h);
    const second = host({});
    reconcileOverlays(current, [{ key: 'a', anchor: at('r0') }], second.h);
    expect(first.handles[0]?.dispose).toHaveBeenCalledTimes(1);
    expect(second.handles).toHaveLength(1);
    reconcileOverlays(current, [{ key: 'a', anchor: at('r0') }], host(null).h);
    expect(second.handles[0]?.dispose).toHaveBeenCalledTimes(1);
    expect(current.size).toBe(0);
  });

  it('key 重复直接抛，与 React 列表同理', () => {
    expect(() =>
      reconcileOverlays(
        new Map(),
        [
          { key: 'a', anchor: at('r0') },
          { key: 'a', anchor: at('r1') },
        ],
        host().h,
      ),
    ).toThrow(/key 重复/);
  });

  it('disposeOverlays 摘光', () => {
    const current = new Map<string | number, OverlayEntry>();
    const { h, handles } = host();
    reconcileOverlays(
      current,
      [
        { key: 'a', anchor: at('r0') },
        { key: 'b', anchor: at('r0') },
      ],
      h,
    );
    disposeOverlays(current);
    expect(handles.every((x) => x.dispose.mock.calls.length === 1)).toBe(true);
    expect(current.size).toBe(0);
  });
});
