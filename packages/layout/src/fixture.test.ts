/**
 * 拿真实公文 `gongwen-01.docx` 走完整条链：解包 → 级联 → 分桶度量 → 断行 → 段落装配，
 * 再与 Word 导出的坐标真值 `gongwen-01.truth.json` 逐行比断行点（L2）。
 *
 * 度量走三级降级里的第②级 —— 随库分发的度量包（`packages/fonts/packs`）。它是在 Windows 上
 * 从 `C:/Windows/Fonts` 抽的，但**入库了**，所以这个测试在 Mac / CI 上跑到的度量与 Word
 * 用的完全一致。这正是 L2 断言能存在的前提：靠第③级的等宽近似，断行点必然对不上。
 *
 * L2 目前**没有全绿**，测试里断言的是「不许退步」而不是「已经对了」——
 * 18 行里对上 11 行（原先是 8 行：空格分桶 + 悬挂半宽两条修完涨上来的）。
 * 剩下的 7 行全从真值第 6 行（0 起）开始连锁：那一行 Word 宁可换行也不肯再挤 14.16pt 的标点，
 * 而我们挤了。这是**临时挤压的上限**没标定，证据表在 `uncalibrated.ts` 的
 * `PUNCT_COMPRESS_RATIO` 里，钉死它要一份专门的 spike 样本。
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
 * 当前 18 行里对上 11 行（含首行与末行）。涨到 11 靠的是两条实测规则：
 * ① 挨着东亚字的空格走 eastAsia 桶（每个空格差 4pt，`items.ts` 的 `applySpaceFont`）；
 * ② 悬挂标点的墨留在版心内、只有空半边吐出去（`HANG_INSIDE_RATIO`）。
 *
 * 第一处分歧在真值第 6 行（行号 0 起）：那一行再多收一个「用」需要挤掉 14.16pt 的标点，行内两个孤立标点
 * 合起来给得起（16pt），Word 却宁可换行。同一份真值里 Word 接受过 9.30pt 的临时挤压、
 * 拒绝过 13.75pt —— 上限就卡在这两个数之间，没有样本能钉死，见 `PUNCT_COMPRESS_RATIO`。
 */
const MIN_L2_MATCH = 11;

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
interface TruthFile {
  pages: { lines: { text: string; x: number; xEnd: number }[] }[];
}

/** 真值的单位是 pt，布局的是 twips（原则 1.3：px 才是禁忌，pt 只在比真值时出现） */
const twipsToPt = (t: number): number => t / 20;

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
    const truth = JSON.parse(readFileSync(TRUTH, 'utf8')) as TruthFile;
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

  it('悬挂的「，」右边缘与真值差 0.05pt —— 这是 L4 级的量，直接检验半宽那条规则', () => {
    // 真值第 4 行以「，」结尾，它的右边缘出版心 8.05pt = 半个三号字，
    // 也就是**墨留在版心内**（`break-class.ts` 的 HANG_INSIDE_RATIO）。
    // 整条规则错了这里会差整整半个字（8pt），远大于 0.5pt 的容差
    const line = laid.flatMap((p) => p.lines).filter((l) => l.fragments.length > 0)[4];
    const last = line?.fragments[line.fragments.length - 1];
    if (last === undefined) throw new Error('第 4 行没有片段');

    const truth = JSON.parse(readFileSync(TRUTH, 'utf8')) as TruthFile;
    const target = truth.pages.flatMap((p) => p.lines)[4];
    const left = doc.resolved.sections[0]?.props.margin.left ?? 0;
    if (target === undefined) throw new Error('真值里没有第 4 行');
    expect(target.text.endsWith('，')).toBe(true);
    expect(twipsToPt(last.x + last.width)).toBeCloseTo(target.xEnd - twipsToPt(left), 1);
  });

  it('第一行与最后一行完全一致 —— 它们不受行内挤压的累积误差影响', () => {
    const ours = ourLines();
    const theirs = truthLines();
    expect(ours[0]).toBe(theirs[0]);
    expect(ours[ours.length - 1]).toBe(theirs[theirs.length - 1]);
  });
});
