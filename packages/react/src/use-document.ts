/**
 * `useDocument(source, options)` —— `UltimateWord.load()` 的 hook 形态。
 *
 * 来源或加载选项变化就重新加载，上一趟还没回来的用 `AbortSignal` 掐掉（只掐得住取字节那一步，
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
  const { fonts, signal, revisions, author } = options;
  const [state, setState] = useState<DocumentState>(() =>
    source === null || source === undefined ? IDLE : { doc: null, loading: true, error: null },
  );

  useEffect(() => {
    if (source === null || source === undefined) {
      setState(IDLE);
      return;
    }
    const controller = new AbortController();
    // 外部取消也必须在没有 AbortSignal.any 的环境生效；本地清理不能反向取消宿主的信号。
    // active 单独区分失效请求与用户取消，否则用户取消后连错误状态也被丢弃，loading 会一直为 true。
    let active = true;
    const abort = () => controller.abort(signal?.reason);
    const detach = () => signal?.removeEventListener('abort', abort);
    if (signal?.aborted) abort();
    else signal?.addEventListener('abort', abort, { once: true });
    setState((prev) =>
      prev.loading && prev.error === null ? prev : { doc: prev.doc, loading: true, error: null },
    );
    const load: LoadOptions = { signal: controller.signal };
    if (fonts !== undefined) load.fonts = fonts;
    // 只依赖字段值：宿主内联 options 不该反复重载，但修订视图与默认署名变化必须生效。
    if (revisions !== undefined) load.revisions = revisions;
    if (author !== undefined) load.author = author;
    UltimateWord.load(source, load).then(
      (doc) => {
        detach();
        if (active) setState({ doc, loading: false, error: null });
      },
      (reason: unknown) => {
        detach();
        if (!active) return;
        setState({
          doc: null,
          loading: false,
          error: reason instanceof Error ? reason : new Error(String(reason)),
        });
      },
    );
    return () => {
      active = false;
      detach();
      controller.abort();
    };
  }, [source, fonts, signal, revisions, author]);

  return state;
}
