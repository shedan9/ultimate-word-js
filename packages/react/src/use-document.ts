/**
 * `useDocument(source, options)` —— `UltimateWord.load()` 的 hook 形态。
 *
 * 来源一换就重新加载，上一趟还没回来的用 `AbortSignal` 掐掉（只掐得住取字节那一步，
 * 排版是同步的 —— `LoadOptions.signal` 的注释说的就是这个）；掐掉的那趟即使回来了也不写状态。
 * `source` 按引用比：字符串 URL 天然稳定，`ArrayBuffer` / `Blob` 要调用方自己 memo，
 * 每次渲染新造一个就会无限重载 —— 与 `useEffect` 的依赖规则一致，不另做深比较。
 */
import { useEffect, useState } from 'react';
import type { LoadOptions, LoadSource, UwDocument } from 'ultimate-word';
import { UltimateWord } from 'ultimate-word';

export interface DocumentState {
  doc: UwDocument | null;
  loading: boolean;
  error: Error | null;
}

const IDLE: DocumentState = { doc: null, loading: false, error: null };

export function useDocument(source: LoadSource | null | undefined, options: LoadOptions = {}): DocumentState {
  const { fonts, signal } = options;
  const [state, setState] = useState<DocumentState>(() =>
    source === null || source === undefined ? IDLE : { doc: null, loading: true, error: null },
  );

  useEffect(() => {
    if (source === null || source === undefined) {
      setState(IDLE);
      return;
    }
    const controller = new AbortController();
    const combined =
      signal === undefined
        ? controller.signal
        : typeof AbortSignal.any === 'function'
          ? AbortSignal.any([controller.signal, signal])
          : controller.signal;
    setState((prev) =>
      prev.loading && prev.error === null ? prev : { doc: prev.doc, loading: true, error: null },
    );
    const load: LoadOptions = { signal: combined };
    if (fonts !== undefined) load.fonts = fonts;
    UltimateWord.load(source, load).then(
      (doc) => {
        if (!controller.signal.aborted) setState({ doc, loading: false, error: null });
      },
      (reason: unknown) => {
        if (controller.signal.aborted) return;
        setState({
          doc: null,
          loading: false,
          error: reason instanceof Error ? reason : new Error(String(reason)),
        });
      },
    );
    return () => controller.abort();
  }, [source, fonts, signal]);

  return state;
}
