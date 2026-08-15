/**
 * 注册表（三级降级）+ 度量包 + `TextMeasurer`。
 *
 * 全部用合成字体，不碰任何系统字体文件 —— 这套测试必须在 macOS / Linux / CI 上
 * 和 Windows 上给出同一个结果，读系统字体会让它变成「在我机器上是绿的」。
 * 真实字体的度量正确性由 metrics.test.ts 的 Word 实测值兜着。
 */
import { createDiagnosticSink, ptToTwips, twipsToPt } from '@uw/core';
import { describe, expect, it } from 'vitest';
import { createTextMeasurer } from './measurer.ts';
import type { RawFontMetrics } from './metrics.ts';
import type { MetricsPack } from './metrics-pack.ts';
import { buildMetricsPack, packAdvance } from './metrics-pack.ts';
import { FontRegistry, fontkitSource, metricsPackSource } from './registry.ts';

/** 宋体家族的形状：unitsPerEm=256，win 跨度恰好 1 em */
const songMetrics: RawFontMetrics = {
  family: 'SimSun',
  postscriptName: 'SimSun',
  unitsPerEm: 256,
  os2: {
    winAscent: 220,
    winDescent: 36,
    typoAscender: 220,
    typoDescender: -36,
    typoLineGap: 36,
    useTypoMetrics: false,
  },
  hhea: { ascender: 220, descender: -36, lineGap: 36 },
};

/** 合成的 fontkit 字体：汉字全角、ASCII 半角、U+2603 没有字形 */
const fakeFont = {
  hasGlyphForCodePoint: (cp: number) => cp !== 0x2603,
  glyphForCodePoint: (cp: number) => ({ advanceWidth: cp < 0x80 ? 128 : 256 }),
};

function pack(): MetricsPack {
  return buildMetricsPack(fakeFont, songMetrics, { family: '宋体' });
}

describe('度量包', () => {
  it('只存与默认宽度不同的码点 —— CJK 字体里例外只有 ASCII 那一小段', () => {
    const p = pack();
    expect(p.defaultAdvance).toBe(256); // 探针 U+4E00 的宽度
    expect(packAdvance(p, 0x4e00)).toBe(256);
    expect(packAdvance(p, 0x41)).toBe(128);
    // 例外表里没有任何全角码点
    expect(Object.keys(p.advances).every((k) => Number(k) < 0x100)).toBe(true);
  });

  it('字体名可以覆盖 —— 文档里写「宋体」，字体自报 SimSun，包要按前者建索引', () => {
    expect(pack().family).toBe('宋体');
    expect(pack().postscriptName).toBe('SimSun');
  });

  it('可结构化克隆 / JSON 往返 —— 它同时是 Worker 传输格式和分发文件', () => {
    const p = pack();
    expect(structuredClone(p)).toEqual(p);
    expect(JSON.parse(JSON.stringify(p))).toEqual(p);
  });

  it('coverage 之外的码点返回 undefined，交给上层降级', () => {
    const p: MetricsPack = { ...pack(), coverage: [[0x4e00, 0x9fff]] };
    expect(packAdvance(p, 0x4e00)).toBe(256);
    expect(packAdvance(p, 0x2603)).toBeUndefined();
    // 例外表优先于 coverage：ASCII 有实测宽度就直接给
    expect(packAdvance(p, 0x41)).toBe(128);
  });
});

describe('注册表的三级降级', () => {
  function registry(): FontRegistry {
    const r = new FontRegistry();
    r.register('SimHei', fontkitSource(fakeFont, songMetrics));
    r.registerPack(pack());
    return r;
  }

  it('按候选名依次查 —— 中文版 Word 写「黑体」，磁盘上叫 SimHei', () => {
    const r = registry();
    expect(r.resolve(['黑体', 'SimHei'])?.family).toBe('SimHei');
    expect(r.resolve(['黑体'])).toBeUndefined();
    expect(r.status(['黑体', 'SimHei'])).toBe('file');
    expect(r.status('黑体')).toBe('missing');
  });

  it('大小写与首尾空白不影响命中', () => {
    expect(registry().resolve([' simhei '])?.family).toBe('SimHei');
  });

  it('度量包注册的字体报 metrics 级', () => {
    expect(registry().status('宋体')).toBe('metrics');
    expect(registry().resolve(['宋体'])?.source.kind).toBe('metrics');
  });

  it('替换表只在原名查不到时才生效 —— 装了真字体就该用真字体', () => {
    const r = registry();
    r.substitute({ 仿宋_GB2312: 'SimHei', 宋体: 'SimHei' });
    expect(r.status('仿宋_GB2312')).toBe('fallback');
    expect(r.resolve(['仿宋_GB2312'])?.family).toBe('SimHei');
    // 宋体自己有度量包，替换表不该把它顶掉
    expect(r.status('宋体')).toBe('metrics');
    expect(r.resolve(['宋体'])?.source.kind).toBe('metrics');
  });

  it('什么都没有是 missing，不是抛异常', () => {
    expect(new FontRegistry().status('方正小标宋简体')).toBe('missing');
    expect(new FontRegistry().resolve(['方正小标宋简体'])).toBeUndefined();
  });

  it('度量包与字体文件两条路给出同一份度量', () => {
    const fromFile = fontkitSource(fakeFont, songMetrics);
    const fromPack = metricsPackSource(pack());
    expect(fromPack.metrics.unitsPerEm).toBe(fromFile.metrics.unitsPerEm);
    expect(fromPack.advance(0x41)).toBe(fromFile.advance(0x41));
    expect(fromPack.advance(0x4e00)).toBe(fromFile.advance(0x4e00));
  });
});

describe('TextMeasurer', () => {
  function measurer(diagnostics = createDiagnosticSink()) {
    const r = new FontRegistry();
    r.register('宋体', fontkitSource(fakeFont, songMetrics));
    return { m: createTextMeasurer(r, { diagnostics }), diagnostics };
  }

  it('宽度是 twips，按字号线性缩放', () => {
    const { m } = measurer();
    const size = ptToTwips(12);
    // 全角 = 1 em = 12pt，半角 = 0.5 em
    expect(twipsToPt(m.advance('宋体', size, 0x4e00))).toBeCloseTo(12, 6);
    expect(twipsToPt(m.advance('宋体', size, 0x41))).toBeCloseTo(6, 6);
  });

  it('批量接口纯数据进出 —— 码点数组进、宽度数组出（为将来抽 WASM 留的形状）', () => {
    const { m } = measurer();
    const cps = Uint32Array.from([0x4e00, 0x41, 0x4e2d]);
    const out = new Float64Array(3);
    m.advances('宋体', ptToTwips(12), cps, out);
    expect([...out]).toEqual([240, 120, 240]);
  });

  it('只量前 count 个，剩下的不碰 —— 断行时按行切片复用同一个缓冲区', () => {
    const { m } = measurer();
    const out = new Float64Array(3).fill(-1);
    m.advances('宋体', ptToTwips(12), Uint32Array.from([0x41, 0x42, 0x43]), out, 2);
    expect([...out]).toEqual([120, 120, -1]);
  });

  it('缓冲区太短直接抛 —— 这是调用方的 bug，不是文档的问题', () => {
    const { m } = measurer();
    expect(() => m.advances('宋体', 240, Uint32Array.from([0x41, 0x42]), new Float64Array(1))).toThrow(
      RangeError,
    );
  });

  it('字体里没有的字形不返回 0 —— 返回 0 会让后面所有字的 x 一起左移', () => {
    const { m } = measurer();
    const w = m.advance('宋体', ptToTwips(12), 0x2603);
    expect(w).toBeGreaterThan(0);
  });

  it('字体缺失记诊断继续跑，且同一款只报一次', () => {
    const { m, diagnostics } = measurer();
    const size = ptToTwips(12);
    m.advance('方正小标宋简体', size, 0x4e00);
    m.advance('方正小标宋简体', size, 0x4e2d);
    m.lineMetrics('方正小标宋简体', size, { eastAsian: true });
    const codes = diagnostics.list().map((d) => d.code);
    expect(codes).toEqual(['font-missing']);
    // 兜底是等宽近似：东亚全角、其余半角
    expect(twipsToPt(m.advance('方正小标宋简体', size, 0x4e00))).toBeCloseTo(12, 6);
    expect(twipsToPt(m.advance('方正小标宋简体', size, 0x41))).toBeCloseTo(6, 6);
  });

  it('行度量走 metrics.ts 的规则，缓存不改变结果', () => {
    const { m } = measurer();
    const size = ptToTwips(12);
    const a = m.lineMetrics('宋体', size, { eastAsian: true });
    const b = m.lineMetrics('宋体', size, { eastAsian: true });
    expect(b).toEqual(a);
    // 东亚：win 跨度 1 em × 1.3
    expect(twipsToPt(a.lineHeight)).toBeCloseTo(15.6, 6);
    // 缓存 key 带 eastAsian，两者不能串
    expect(m.lineMetrics('宋体', size, { eastAsian: false }).lineHeight).not.toBe(a.lineHeight);
  });

  it('缓存超限按最久未用淘汰，结果依然正确', () => {
    const r = new FontRegistry();
    r.register('宋体', fontkitSource(fakeFont, songMetrics));
    const m = createTextMeasurer(r, { lineMetricsCacheSize: 2 });
    const h = (pt: number) => m.lineMetrics('宋体', ptToTwips(pt), { eastAsian: true }).lineHeight;
    const first = h(12);
    h(14);
    h(16);
    expect(h(12)).toBe(first);
  });

  it('候选名展开由调用方注入 —— fonts 不认识 model 的 fontTable', () => {
    const r = new FontRegistry();
    r.register('SimHei', fontkitSource(fakeFont, songMetrics));
    const m = createTextMeasurer(r, { candidates: (f) => (f === '黑体' ? ['黑体', 'SimHei'] : [f]) });
    expect(m.status('黑体')).toBe('file');
    expect(twipsToPt(m.advance('黑体', ptToTwips(12), 0x4e00))).toBeCloseTo(12, 6);
  });
});
