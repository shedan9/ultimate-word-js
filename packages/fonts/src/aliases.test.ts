/**
 * 同一款字体的中英文名互通（aliases.ts）。拿随库的度量包验：它们登记的是中文名，
 * 文档里写英文名（WPS、英文版 Word）照样要命中，且命中的就是那一份度量。
 */
import { createDiagnosticSink } from '@uw/core';
import { describe, expect, it } from 'vitest';
import { createTextMeasurer } from './measurer.ts';
import { bundledPacks } from './packs.ts';
import type { FontSource } from './registry.ts';
import { FontRegistry } from './registry.ts';

function registry(): FontRegistry {
  const r = new FontRegistry();
  for (const pack of bundledPacks()) r.registerMetrics(pack);
  return r;
}

describe('中英文字体名互通', () => {
  it('英文名命中随库的中文名度量包，命中的是同一份', () => {
    const r = registry();
    for (const [en, zh] of [
      ['SimSun', '宋体'],
      ['simhei', '黑体'],
      ['FangSong', '仿宋'],
      ['KaiTi', '楷体'],
      ['DengXian', '等线'],
      ['Microsoft YaHei', '微软雅黑'],
    ] as const) {
      expect(r.resolve([en])?.source, en).toBe(r.resolve([zh])?.source);
      expect(r.status(en)).toBe('metrics');
    }
  });

  it('反方向同样成立：宿主按英文名注册的字体文件，文档写中文名也找得到', () => {
    const r = new FontRegistry();
    const source = { kind: 'metrics' } as unknown as FontSource;
    r.register('SimHei', source);
    expect(r.resolve(['黑体'])?.source).toBe(source);
  });

  it('直接命中优先于别名：宿主真注册了叫 SimSun 的字体就用它自己', () => {
    const r = registry();
    const own = { kind: 'file' } as unknown as FontSource;
    r.register('SimSun', own);
    expect(r.resolve(['SimSun'])?.source).toBe(own);
    expect(r.resolve(['宋体'])?.source).not.toBe(own);
  });

  it('度量器不再为英文名报 font-missing', () => {
    const sink = createDiagnosticSink();
    const m = createTextMeasurer(registry(), { diagnostics: sink });
    m.eastAsianFont('SimSun');
    expect(sink.list().filter((d) => d.code === 'font-missing')).toEqual([]);
    expect(m.eastAsianFont('SimSun')).toBe(true);
  });
});
