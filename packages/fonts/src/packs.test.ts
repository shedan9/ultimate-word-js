import { describe, expect, it } from 'vitest';
import { loadBundledPacks } from './load-node.ts';
import { bundledPacks } from './packs.ts';

describe('bundledPacks（无 fs 入口）', () => {
  it('与 packs/index.json 列的那批一模一样 —— 手写的 import 列表不许漂', () => {
    const fromFs = loadBundledPacks()
      .map((p) => p.family)
      .sort();
    const fromImport = bundledPacks()
      .map((p) => p.family)
      .sort();
    expect(fromImport).toEqual(fromFs);
  });

  it('每个包的内容与 fs 读到的一致', () => {
    const byFamily = new Map(loadBundledPacks().map((p) => [p.family, p]));
    for (const p of bundledPacks()) expect(p).toEqual(byFamily.get(p.family));
  });
});
