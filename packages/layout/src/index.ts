/**
 * 布局引擎。Phase 2 填充。
 *
 * 约束（从第一天就守住，否则后面搬不动）：
 * - 本包**不得** import 任何 DOM API，度量能力靠注入的接口传进来 ——
 *   这样才能在 Worker / Node 里跑，也是未来把 linebreak + measure 换成 Rust→WASM 的前提
 * - linebreak / measure 的接口保持纯数据进出（码点数组 + 宽度数组 → 断点数组），
 *   不持有任何 JS 对象引用
 */
export const PACKAGE_NAME = '@uw/layout';
