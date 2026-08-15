/**
 * 诊断 —— 「文档还能看，只是某处不对」的记录方式。
 *
 * 与 `UwError` 的分界见 [errors.ts](./errors.ts)。诊断是**纯数据**（原则 1.1），
 * 因为它要跟着 `LayoutResult` 一起过 Worker 边界，最终从 `doc.diagnostics` 暴露出去。
 */

export type DiagnosticSeverity = 'warn' | 'info';

export interface Diagnostic {
  severity: DiagnosticSeverity;
  /** 稳定的短码，供调用方分流 / 上报聚合，如 `unknown-element` */
  code: string;
  /** 给人看的说明，中文 */
  message: string;
  /** 出处：部件名 + 元素路径，能定位到就带上 */
  part?: string;
  path?: string;
}

/**
 * 诊断收集器。
 *
 * 用普通函数闭包而不是类：它会被一路传进解析器的各层，闭包形态不会诱使谁去继承它，
 * 也不会有人不小心把它塞进纯数据输出里（`list()` 返回的才是能过结构化克隆的东西）。
 */
export interface DiagnosticSink {
  warn(code: string, message: string, where?: { part?: string; path?: string }): void;
  info(code: string, message: string, where?: { part?: string; path?: string }): void;
  /** 快照，调用方拿到的是副本 */
  list(): Diagnostic[];
}

export function createDiagnosticSink(): DiagnosticSink {
  const items: Diagnostic[] = [];
  const push =
    (severity: DiagnosticSeverity) =>
    (code: string, message: string, where: { part?: string; path?: string } = {}): void => {
      const d: Diagnostic = { severity, code, message };
      // exactOptionalPropertyTypes：可选字段要么不写，要么给确定值，不能写 undefined
      if (where.part !== undefined) d.part = where.part;
      if (where.path !== undefined) d.path = where.path;
      items.push(d);
    };
  return {
    warn: push('warn'),
    info: push('info'),
    list: () => items.slice(),
  };
}
