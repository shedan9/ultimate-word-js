import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const pExecFile = promisify(execFile);

let cached: string | undefined;

/** 优先 pwsh 7，回落 Windows PowerShell 5.1 —— 两者跑 Word COM 都可以 */
export async function powershellExe(): Promise<string> {
  if (cached) return cached;
  for (const exe of ['pwsh', 'powershell']) {
    try {
      await pExecFile(exe, ['-NoProfile', '-NonInteractive', '-Command', '$PSVersionTable.PSVersion.Major']);
      cached = exe;
      return exe;
    } catch {
      // 试下一个
    }
  }
  throw new Error('找不到 pwsh 或 powershell —— 真值流水线需要 Windows + 已安装的 Word');
}

export async function runScript(script: string, args: string[]): Promise<string> {
  const exe = await powershellExe();
  const { stdout } = await pExecFile(
    exe,
    ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, ...args],
    { windowsHide: true, maxBuffer: 16 * 1024 * 1024 },
  );
  return stdout.trim();
}
