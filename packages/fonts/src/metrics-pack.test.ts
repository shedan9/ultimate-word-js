/**
 * 度量包：格式本身 + 随库入库的那 17 份。
 *
 * 入库的包是**在 Windows 上从 `C:/Windows/Fonts` 抽的**（`tools/build-packs.ts`），
 * 这个测试跑在所有平台上 —— 那正是包存在的理由。这里断言的是「包还在、还能用、
 * 关键数值没变」，因为下游的坐标级真值断言（L2/L4）全站在它们上面：
 * 包里的宽度错一点，断行点就换一个字，而那种错很难反查到这一层。
 */
import { describe, expect, it } from 'vitest';
import { loadBundledPacks } from './load-node.ts';
import type { RawFontMetrics } from './metrics.ts';
import type { GlyphSource } from './metrics-pack.ts';
import { buildMetricsPack, packAdvance, symbolSampleCodePoints } from './metrics-pack.ts';
import { FontRegistry } from './registry.ts';

const metrics: RawFontMetrics = {
  family: 'Test',
  postscriptName: 'Test',
  unitsPerEm: 1000,
  os2: {
    winAscent: 800,
    winDescent: 200,
    typoAscender: 800,
    typoDescender: -200,
    typoLineGap: 0,
    useTypoMetrics: false,
  },
  hhea: { ascender: 800, descender: -200, lineGap: 0 },
};

/** 造一款假字体：给定「码点 → 宽度」，没列出来的码点就是没有字形 */
function fakeFont(widths: Record<number, number>): GlyphSource {
  return {
    hasGlyphForCodePoint: (cp) => cp in widths,
    glyphForCodePoint: (cp) => {
      const w = widths[cp];
      return w === undefined ? null : { advanceWidth: w };
    },
    characterSet: Object.keys(widths).map(Number),
  };
}

describe('度量包格式', () => {
  it('覆盖码点压成升序区间 —— 宋体的 28850 个码点靠这个从 200KB 降到 2KB', () => {
    const pack = buildMetricsPack(fakeFont({ 32: 500, 33: 500, 34: 500, 100: 500 }), metrics, {
      sample: [32, 33, 34, 100],
    });
    expect(pack.coverage).toEqual([
      [32, 34],
      [100, 100],
    ]);
  });

  it('覆盖范围外的码点返回 undefined，让上层知道「这款字体没这个字」', () => {
    const pack = buildMetricsPack(fakeFont({ 32: 500, 33: 500 }), metrics, { sample: [32, 33] });
    expect(packAdvance(pack, 32)).toBe(500);
    expect(packAdvance(pack, 0x4e00)).toBeUndefined();
  });

  it('CJK 字体的默认宽度取探针字形（U+4E00），于是汉字一个例外都不用存', () => {
    const pack = buildMetricsPack(fakeFont({ 19968: 1000, 19969: 1000, 32: 500 }), metrics, {
      sample: [32, 0x4e00, 0x4e01],
    });
    expect(pack.defaultAdvance).toBe(1000);
    expect(Object.keys(pack.advances)).toEqual(['32']);
  });

  it('拉丁字体没有探针字形时默认宽度取众数，而不是退成一个全角宽', () => {
    // 退成 unitsPerEm 就等于宣布「没列出来的字都是全角」，希腊字母、西里尔字母全被算成汉字宽
    const pack = buildMetricsPack(fakeFont({ 65: 600, 66: 600, 67: 600, 105: 300 }), metrics, {
      sample: [65, 66, 67, 105],
    });
    expect(pack.defaultAdvance).toBe(600);
    expect(pack.advances).toEqual({ '105': 300 });
  });
});

describe('随库度量包', () => {
  const packs = loadBundledPacks();
  const byFamily = new Map(packs.map((p) => [p.family, p]));

  it('公文那四款中文字体在册，且按**文档里写的中文名**建索引', () => {
    // 中文版 Word 在 w:rFonts 里写的就是「仿宋」，磁盘上的文件才叫 simfang.ttf。
    // 按 PostScript 名建索引会让主路径整个落空，而落空是静默的
    for (const family of ['宋体', '仿宋', '黑体', '楷体']) {
      expect(byFamily.get(family)?.version).toBe(1);
    }
  });

  it('中文字体的汉字精确等宽 1 em —— 「一行 28 字」的公文版心靠这一条', () => {
    for (const family of ['宋体', '仿宋', '黑体', '楷体']) {
      const pack = byFamily.get(family);
      if (pack === undefined) throw new Error(`度量包缺失：${family}`);
      expect(pack.defaultAdvance).toBe(pack.unitsPerEm);
      expect(packAdvance(pack, 0x4e00)).toBe(pack.unitsPerEm);
    }
  });

  it('仿宋的 win 度量与 Phase 0 标定的一致 —— 行高全靠它', () => {
    const pack = byFamily.get('仿宋');
    expect(pack?.unitsPerEm).toBe(256);
    expect(pack?.os2.winAscent).toBe(220);
    expect(pack?.os2.winDescent).toBe(36);
  });

  it('等线与微软雅黑的 win 跨度明显不是 1 em —— 缺了它们的包，行高会错三成', () => {
    const deng = byFamily.get('等线');
    const yahei = byFamily.get('微软雅黑');
    const span = (p: typeof deng): number =>
      p === undefined ? 0 : (p.os2.winAscent + p.os2.winDescent) / p.unitsPerEm;
    expect(span(deng)).toBeCloseTo(1.042, 3);
    expect(span(yahei)).toBeCloseTo(1.3198, 3);
  });

  it('符号字体的码点在 0x00–0xFF —— docx 里写的 F0B7 要先减掉 0xF000', () => {
    const symbol = byFamily.get('Symbol');
    if (symbol === undefined) throw new Error('度量包缺失：Symbol');
    expect(symbol.coverage).toEqual([[0, 255]]);
    // numbering.xml 里实心圆点就是 Symbol 的 0xB7，宽度必须查得到
    expect(packAdvance(symbol, 0xb7)).toBeGreaterThan(0);
    // 私用区那个码点查不到，正是「减 0xF000 归调用方」这条约定的证据
    expect(packAdvance(symbol, 0xf0b7)).toBeUndefined();
    expect(symbolSampleCodePoints()).toHaveLength(256);
  });

  it('拉丁字体的默认宽度不是全角 —— 众数那一档生效了', () => {
    for (const family of ['Times New Roman', 'Arial', 'Georgia']) {
      const pack = byFamily.get(family);
      expect(pack?.defaultAdvance).toBeLessThan(pack?.unitsPerEm ?? 0);
    }
  });

  it('注册进注册表之后，字体状态是第②级而不是兜底', () => {
    const registry = new FontRegistry();
    for (const pack of packs) registry.registerMetrics(pack);
    expect(registry.status('仿宋')).toBe('metrics');
    expect(registry.status('Times New Roman')).toBe('metrics');
    // 没抽的字体照旧报 missing —— 包不是万能的，缺了就该被看见
    expect(registry.status('方正小标宋简体')).toBe('missing');
  });

  it('全部可结构化克隆 —— 度量包同时是 Worker 边界上的传输格式', () => {
    expect(structuredClone(packs)).toEqual(packs);
  });
});
