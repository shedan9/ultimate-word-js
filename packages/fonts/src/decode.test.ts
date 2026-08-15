/**
 * 字节 → 字体。
 *
 * 这里 mock 掉 fontkit：要测的是**我们和 fontkit 之间的那层契约**（喂什么类型的字节、
 * 字体集怎么拆包），不是 fontkit 自己解析 TrueType 的能力。真字体文件不能进这套测试 ——
 * 仓库里没有可分发的字体，读系统字体又会让 macOS / Linux / Windows 三边结果不一致。
 */
import { beforeEach, describe, expect, it, vi } from 'vitest';

const create = vi.fn();
vi.mock('fontkit', () => ({ create: (...args: unknown[]) => create(...args) }));

const { decodeFont, fontSourceFromBytes, unwrapFont } = await import('./decode.ts');

/** 最小可用的假字体：readRawMetrics 只要 OS/2 与 hhea */
function fakeFont(postscriptName: string) {
  return {
    familyName: postscriptName,
    postscriptName,
    unitsPerEm: 1000,
    'OS/2': {
      winAscent: 800,
      winDescent: 200,
      typoAscender: 800,
      typoDescender: -200,
      typoLineGap: 0,
      fsSelection: 0,
    },
    hhea: { ascent: 800, descent: -200, lineGap: 0 },
    hasGlyphForCodePoint: () => true,
    glyphForCodePoint: () => ({ advanceWidth: 500 }),
  };
}

beforeEach(() => {
  create.mockReset();
});

describe('字节归一化', () => {
  it('ArrayBuffer 要转成 Uint8Array 再交给 fontkit —— 直接喂它会在 DataView 处抛', () => {
    // 这不是洁癖：fetch(...).then(r => r.arrayBuffer()) 拿到的正是 ArrayBuffer，
    // 而 fontkit 只吃 Uint8Array。转换放在库里，调用方不必知道这层区别
    create.mockReturnValue(fakeFont('X'));
    decodeFont(new ArrayBuffer(8));
    expect(create.mock.calls[0]?.[0]).toBeInstanceOf(Uint8Array);
  });

  it('已经是 Uint8Array 就原样传，不复制一份', () => {
    create.mockReturnValue(fakeFont('X'));
    const bytes = new Uint8Array(8);
    decodeFont(bytes);
    expect(create.mock.calls[0]?.[0]).toBe(bytes);
  });

  it('postscriptName 透传给 fontkit', () => {
    create.mockReturnValue(fakeFont('SimSun'));
    decodeFont(new Uint8Array(8), 'SimSun');
    expect(create.mock.calls[0]?.[1]).toBe('SimSun');
  });
});

describe('字体集拆包', () => {
  it('按 postscriptName 取其中一款 —— simsun.ttc 里同时装着 SimSun 与 NSimSun', () => {
    const sim = fakeFont('SimSun');
    const collection = {
      fonts: [fakeFont('NSimSun'), sim],
      getFont: (n: string) => (n === 'SimSun' ? sim : null),
    };
    expect(unwrapFont(collection, 'SimSun', 'x.ttc')).toBe(sim);
  });

  it('不指定就取第一款', () => {
    const first = fakeFont('A');
    expect(unwrapFont({ fonts: [first, fakeFont('B')] }, undefined, 'x.ttc')).toBe(first);
  });

  it('集合里没有指定的那款就抛 —— 拿错字体量出来的宽度是错的，静默降级更糟', () => {
    const collection = { fonts: [fakeFont('A')], getFont: () => null };
    expect(() => unwrapFont(collection, '不存在', 'x.ttc')).toThrow('x.ttc 里没有 不存在');
    expect(() => unwrapFont({ fonts: [] }, undefined, 'x.ttc')).toThrow('空的字体集');
  });

  it('单款字体原样返回', () => {
    const font = fakeFont('A');
    expect(unwrapFont(font, undefined, 'x.ttf')).toBe(font);
  });
});

describe('字节 → FontSource', () => {
  it('产出降级链第 ①级，可直接注册进 FontRegistry', () => {
    create.mockReturnValue(fakeFont('SimHei'));
    const src = fontSourceFromBytes(new ArrayBuffer(8));
    expect(src.kind).toBe('file');
    expect(src.metrics.unitsPerEm).toBe(1000);
    expect(src.metrics.os2.winAscent).toBe(800);
    expect(src.advance(0x41)).toBe(500);
  });
});
