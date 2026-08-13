/**
 * 平台守卫。
 *
 * 真值的**生成**链路绑死 Windows：导出 PDF 要 Word COM，行高穿刺要读
 * `C:/Windows/Fonts` 下的真实字体。真值的**消费**（跑测试、写布局代码）则完全跨平台 ——
 * `*.truth.json` 已经入库。
 *
 * 所以在 Mac / Linux 上跑这两个入口不是 bug，是用错了工具；报错信息要说清楚这一点，
 * 而不是甩一个 ENOENT: C:/Windows/Fonts/simfang.ttf 让人以为是路径写错了。
 */

export interface WindowsOnlyReason {
  /** 这个入口在做什么 */
  tool: string;
  /** 依赖 Windows 的具体原因 */
  needs: string;
}

/** 非 Windows 平台直接退出（退出码 2 = 环境不满足，区别于 1 = 断言失败） */
export function assertWindows({ tool, needs }: WindowsOnlyReason): void {
  if (process.platform === 'win32') return;
  console.error(
    [
      `${tool} 只能在 Windows 上跑：${needs}。`,
      `当前平台是 ${process.platform}。`,
      '',
      '真值的生成需要 Windows + Word，但真值的「使用」不需要 ——',
      'apps/fidelity/fixtures/*.truth.json 已经入库，直接读就行：',
      '  pnpm turbo run test          # 跨平台，绿的',
      '  pnpm --filter @uw/playground dev',
      '',
      '新增 fixture 或改动 fixture spec 时，回 Windows 跑一次 `pnpm truth` 重新生成即可。',
    ].join('\n'),
  );
  process.exit(2);
}
