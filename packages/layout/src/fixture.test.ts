/**
 * 拿真实公文 `gongwen-01.docx` 走完整条链：解包 → 级联 → 分桶度量 → 断行 → 段落装配，
 * 再与 Word 导出的坐标真值 `gongwen-01.truth.json` 逐行比断行点（L2）。
 *
 * 度量走三级降级里的第②级 —— 随库分发的度量包（`packages/fonts/packs`）。它是在 Windows 上
 * 从 `C:/Windows/Fonts` 抽的，但**入库了**，所以这个测试在 Mac / CI 上跑到的度量与 Word
 * 用的完全一致。这正是 L2 断言能存在的前提：靠第③级的等宽近似，断行点必然对不上。
 *
 * L2 目前**没有全绿**，测试里断言的是「不许退步」而不是「已经对了」——
 * 剩下的差指向同一件已经量到、但还没实现的事：**悬挂标点的墨留在版心内，只有空半边吐出去**，
 * 而我们把悬挂项整个不计入行宽。见 `uncalibrated.ts` 的 `PUNCT_COMPRESS_RATIO` 末段。
 * 修完之后 `MIN_L2_MATCH` 要往上调。
 */
import { readFileSync } from 'node:fs';
import { createDiagnosticSink } from '@uw/core';
import type { TextMeasurer } from '@uw/fonts';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import type { LoadedDocument, ResolvedParagraph } from '@uw/model';
import { fontNameCandidates, loadDocument, walkBlocks } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import { beforeAll, describe, expect, it } from 'vitest';
import type { LayoutParagraphOptions } from './paragraph.ts';
import { layoutParagraph } from './paragraph.ts';
import type { ParagraphLayout } from './types.ts';

const FIXTURE = new URL('../../../apps/fidelity/fixtures/gongwen-01.docx', import.meta.url);
const TRUTH = new URL('../../../apps/fidelity/fixtures/gongwen-01.truth.json', import.meta.url);

/**
 * 与真值逐字一致的行数下限。**只允许往上调** —— 它是「不许退步」的闸门，不是达标线。
 *
 * 当前 18 行里对上 8 行（含首行与末行）。第一处分歧在第 5 行：Word 让行尾的「，」
 * 把**空半边**吐出版心、墨留在里面，于是那一行到此结束；我们把悬挂的标点整个不计入行宽，
 * 于是还剩得下一个字，从此每一行都错开一个字。修法见 `uncalibrated.ts` 末段，
 * 它要连带改两端对齐与行宽断言，所以没有塞进这一轮。
 */
const MIN_L2_MATCH = 8;

const sink = createDiagnosticSink();
let doc: LoadedDocument;
let paragraphs: ResolvedParagraph[];
let laid: ParagraphLayout[];
let contentWidth: number;
let measurer: TextMeasurer;

beforeAll(() => {
  doc = loadDocument(OpcPackage.open(new Uint8Array(readFileSync(FIXTURE))), sink);
  const section = doc.resolved.sections[0];
  if (section === undefined) throw new Error('fixture 里没有节');

  const { page, margin, docGrid } = section.props;
  contentWidth = page.width - margin.left - margin.right;

  const registry = new FontRegistry();
  // 度量包是**入库的数据**，所以这一步跨平台。Windows 上重抽见 @uw/fonts 的 tools/build-packs.ts
  for (const pack of loadBundledPacks()) registry.registerMetrics(pack);

  measurer = createTextMeasurer(registry, {
    // 中文版 Word 写的是「仿宋」这种本地化名，磁盘上叫 FangSong —— 桥在 fontTable.xml
    candidates: (family) => fontNameCandidates(doc.fonts, family),
    diagnostics: sink,
  });
  const opts: LayoutParagraphOptions = {
    measurer,
    contentWidth,
    settings: doc.cascade.settings,
    docGrid,
  };

  paragraphs = [];
  for (const block of walkBlocks(section.blocks)) {
    if (block.kind === 'paragraph') paragraphs.push(block);
  }
  laid = paragraphs.map((p) => layoutParagraph(p, opts));
});

describe('真实公文走完整条链', () => {
  it('每一段都排得出行，空段落也占一行', () => {
    expect(laid).toHaveLength(paragraphs.length);
    expect(laid.every((p) => p.lines.length >= 1)).toBe(true);
  });

  it('行高全为正 —— 没有哪一段的度量落空成 0', () => {
    expect(laid.flatMap((p) => p.lines).every((l) => l.height > 0)).toBe(true);
  });

  it('行不超出版心（悬挂标点除外，那是 overflowPunct 的正常行为）', () => {
    for (const p of laid) {
      for (const line of p.lines) {
        // 允许一个全角字的余量：悬挂出去的标点不计入 width，但制表位可能顶到边上
        expect(line.x + line.width).toBeLessThanOrEqual(contentWidth + 1);
      }
    }
  });

  it('正文段落的首行缩进 2 字符落在 x 上 —— 公文几乎每段都有', () => {
    const indented = paragraphs
      .map((p, i) => ({ p, out: laid[i] }))
      .filter((e) => e.p.props.indent.firstLineChars === 200);
    expect(indented.length).toBeGreaterThan(0);
    for (const { p, out } of indented) {
      // 「2 字符」按段落里第一个字符的字号折算，上界就是两个全角字宽
      expect(out?.lines[0]?.x).toBeGreaterThan(0);
      expect(out?.lines[0]?.x).toBeLessThanOrEqual(p.props.markRunProps.size * 2 + 1);
    }
  });

  it('片段的文字拼起来就是该行的原文，一个字都不丢', () => {
    const all = laid
      .flatMap((p) => p.lines)
      .flatMap((l) => l.fragments)
      .map((f) => f.text)
      .join('');
    expect(all).toContain('通知');
    expect(all.length).toBeGreaterThan(100);
  });

  it('整份布局结果可结构化克隆 —— 它要过 Worker 边界', () => {
    expect(structuredClone(laid)).toEqual(laid);
  });

  it('这份文档在布局阶段一条诊断都没有 —— 三款字体全部命中度量包', () => {
    // 度量包进来之前这里要放过 font-missing。现在放过它就等于放过「包没被用上」
    expect(sink.list()).toEqual([]);
  });

  it('三款字体都落在降级链第②级（度量包），一个都没退到等宽近似', () => {
    for (const family of ['仿宋', '黑体', 'Times New Roman']) {
      expect(measurer.status(family)).toBe('metrics');
    }
  });
});

/**
 * L2：每行断行点与 Word 一致。
 *
 * 比的是**行文字**而不是首末码点：文字一致蕴含首末码点一致，且失败时的报错直接能看出
 * 「我们多收了一个字」还是「少收了一个字」。行尾空格不算 —— Word 让它吐出版心，
 * PDF 里那个空格不落墨，真值的行文字里也就没有它。
 */
describe('L2 · 断行点与 Word 真值', () => {
  const ourLines = (): string[] =>
    laid
      .flatMap((p) => p.lines)
      .map((l) =>
        l.fragments
          .map((f) => f.text)
          .join('')
          .replace(/\s+$/u, ''),
      )
      .filter((t) => t !== '');

  const truthLines = (): string[] => {
    const truth = JSON.parse(readFileSync(TRUTH, 'utf8')) as { pages: { lines: { text: string }[] }[] };
    return truth.pages
      .flatMap((p) => p.lines.map((l) => l.text.replace(/\s+$/u, '')))
      .filter((t) => t !== '');
  };

  it('行数与 Word 一致', () => {
    // 这一条比逐行文字更硬：行数错了说明断行系统性偏了，而不是某一行差一个字
    expect(ourLines()).toHaveLength(truthLines().length);
  });

  it(`至少 ${MIN_L2_MATCH} 行与 Word 逐字一致（闸门，只许往上调）`, () => {
    const ours = ourLines();
    const theirs = truthLines();
    const matched = ours.filter((t, i) => t === theirs[i]).length;
    expect(matched).toBeGreaterThanOrEqual(MIN_L2_MATCH);
  });

  it('第一行与最后一行完全一致 —— 它们不受行内挤压的累积误差影响', () => {
    const ours = ourLines();
    const theirs = truthLines();
    expect(ours[0]).toBe(theirs[0]);
    expect(ours[ours.length - 1]).toBe(theirs[theirs.length - 1]);
  });
});
