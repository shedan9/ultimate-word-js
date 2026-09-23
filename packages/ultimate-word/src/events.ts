/**
 * 事件（api.md §11）。`on()` 返回 `Disposable`，一个事件可以挂任意多个监听者 ——
 * 属性式的 `onXxx` 天然只能挂一个，这是 api.md 选 `名词:动词` 的原因。
 *
 * 监听者抛错**不传回触发方**：事件是在事务提交、IntersectionObserver 回调这类内部时机派发的，
 * 一个宿主监听者的 bug 不该让事务半途而废（模型已经提交了，视图却没刷新），也不该挡住
 * 排在它后面的监听者。错误交给 `reportError`（浏览器 / Node 都有），宿主的全局错误处理照样看得见。
 */
import type { Disposable } from './view.ts';

type Listener<T> = (payload: T) => void;

export class Emitter<M extends object> {
  readonly #listeners = new Map<keyof M, Set<Listener<never>>>();

  on<K extends keyof M>(type: K, listener: Listener<M[K]>): Disposable {
    let bucket = this.#listeners.get(type);
    if (bucket === undefined) {
      bucket = new Set();
      this.#listeners.set(type, bucket);
    }
    // 同一个函数挂两次算两个监听者，各自 dispose —— Set 会把它们并成一个，所以包一层
    const entry: Listener<M[K]> = (payload) => listener(payload);
    bucket.add(entry as Listener<never>);
    return {
      dispose: () => {
        bucket.delete(entry as Listener<never>);
      },
    };
  }

  /** 有没有人在听。算 payload 要花力气（排序、查表）的事件先问一句 */
  has(type: keyof M): boolean {
    return (this.#listeners.get(type)?.size ?? 0) > 0;
  }

  emit<K extends keyof M>(type: K, payload: M[K]): void {
    const bucket = this.#listeners.get(type);
    if (bucket === undefined) return;
    // 拷一份再遍历：监听者在回调里 dispose 自己或别人，不影响这一轮
    for (const listener of [...bucket]) {
      try {
        (listener as Listener<M[K]>)(payload);
      } catch (error) {
        report(error);
      }
    }
  }

  clear(): void {
    this.#listeners.clear();
  }
}

function report(error: unknown): void {
  if (typeof globalThis.reportError === 'function') globalThis.reportError(error);
  else
    queueMicrotask(() => {
      throw error;
    });
}
