/**
 * zip 解包。fflate 的一层薄封装，唯一的价值是把它的异常翻译成 `UwError`。
 *
 * 用同步 API：docx 通常几百 KB 到几 MB，解压是毫秒级；异步版会把 `load()` 的调用链
 * 整个染成异步，而真正值得异步化的是**整个布局跑进 Worker**，不是这一步。
 */
import { UwError, UwErrorCode } from '@uw/core';
import { unzipSync } from 'fflate';

/** zip 条目名（不带前导斜杠）→ 字节 */
export type ZipEntries = Map<string, Uint8Array>;

export function unzip(data: Uint8Array): ZipEntries {
  let raw: Record<string, Uint8Array>;
  try {
    raw = unzipSync(data);
  } catch (cause) {
    throw new UwError(UwErrorCode.NOT_A_ZIP, '无法解包：这个文件不是有效的 zip（.docx 本质是 zip）', {
      cause,
    });
  }
  const entries: ZipEntries = new Map();
  for (const [name, bytes] of Object.entries(raw)) {
    // 目录条目（zip 里以 / 结尾的空条目）不是部件，跳过
    if (name.endsWith('/')) continue;
    entries.set(name, bytes);
  }
  return entries;
}
