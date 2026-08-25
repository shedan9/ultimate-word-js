/**
 * 图片几何的真值回归：`spike-image-01/02` 两份样本，**逐行 + 逐图**与 Word 对。
 *
 * 与别的 fixture 测试有两处差别，都是被样本逼出来的：
 *
 * ① **行比的是逐行增量，不是累加的绝对 y**。Word 自己的行位置带着 ±0.12pt 的抖动
 *    （01 里那几行纯文字参照行，行距在 15.48–15.62pt 之间跳），44 行图叠起来能累到 1.5pt，
 *    早已越过 L3 判据 —— 那是 Word 内部取整的锅，不是行盒规则的锅。而行盒规则决定的
 *    恰好是**增量**：一行有多高、一张图把行撑高多少。每页第一行仍按绝对 y 比。
 * ② **内嵌图比的是它相对本行基线的抬升**，浮动图才比纸坐标。理由同 ①。
 *
 * 规则怎么标定出来的见 `apps/fidelity` 的 `spike:image`（8 种组合逐行逐图跑，唯一满分）。
 * 跨平台：docx 与 truth.json 都入库，度量走随库的度量包，所以 CI 上也跑得了。
 */
import { readFileSync } from 'node:fs';
import { createDiagnosticSink, twipsToPt } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { layoutDocumentWithFields } from './fields.ts';
import type { DocumentLayout, PageLayout } from './page.ts';

const FIXTURES = new URL('../../../apps/fidelity/fixtures/', import.meta.url);
/** L3 / L4 的判据 */
const TOLERANCE_PT = 0.5;

interface TruthImage {
  x: number;
  y: number;
  w: number;
  h: number;
  yBottom: number;
}
interface Truth {
  pageCount: number;
  pages: { lines: { y: number }[]; images?: TruthImage[] }[];
}

interface Rect {
  x: number;
  y: number;
}

function layout(name: string): { doc: DocumentLayout; truth: Truth; diagnostics: number } {
  const sink = createDiagnosticSink();
  const bytes = new Uint8Array(readFileSync(new URL(`${name}.docx`, FIXTURES)));
  const loaded = loadDocument(OpcPackage.open(bytes), sink);
  const registry = new FontRegistry();
  for (const pack of loadBundledPacks()) registry.registerMetrics(pack);
  const measurer = createTextMeasurer(registry, {
    candidates: (family) => fontNameCandidates(loaded.fonts, family),
    diagnostics: sink,
  });
  return {
    doc: layoutDocumentWithFields(loaded.resolved, loaded.fields, {
      measurer,
      settings: loaded.cascade.settings,
      headerFooters: loaded.headerFooters,
    }).layout,
    truth: JSON.parse(readFileSync(new URL(`${name}.truth.json`, FIXTURES), 'utf8')) as Truth,
    diagnostics: sink.list().length,
  };
}

/** 一页上每一行的绝对基线（pt）。空段落在 PDF 里不落墨，真值没有它，这边也要滤掉 */
function baselines(page: PageLayout): number[] {
  const out: number[] = [];
  for (const b of page.blocks) {
    if (b.kind !== 'paragraph') continue;
    for (const placed of b.lines) {
      const text = placed.line.fragments.map((f) => f.text).join('');
      if (text.trim() === '') continue;
      out.push(twipsToPt(page.geometry.content.y + placed.y + placed.line.baseline));
    }
  }
  return out.sort((a, b) => a - b);
}

/** 一页上每一张图：内嵌的从行里算（y = 基线 − 高 − 抬升），浮动的分页那一步已经算好了 */
function rects(page: PageLayout): { x: number; y: number; h: number }[] {
  const out: { x: number; y: number; h: number }[] = [];
  const g = page.geometry.content;
  for (const b of page.blocks) {
    if (b.kind !== 'paragraph') continue;
    for (const placed of b.lines) {
      for (const obj of placed.line.objects ?? []) {
        const bottom = g.y + placed.y + placed.line.baseline - (obj.raise ?? 0);
        out.push({
          x: twipsToPt(g.x + obj.x),
          y: twipsToPt(bottom - obj.height),
          h: twipsToPt(obj.height),
        });
      }
    }
  }
  for (const f of page.floats ?? []) {
    out.push({ x: twipsToPt(f.x), y: twipsToPt(f.y), h: twipsToPt(f.height) });
  }
  return out.sort((a, b) => a.y - b.y || a.x - b.x);
}

/** 内嵌图换成「x + 相对本行基线的抬升」，浮动图原样用纸坐标（理由见文件头 ②） */
function comparable(
  rs: readonly { x: number; y: number; h: number }[],
  lines: readonly number[],
  mode: 'inline' | 'float',
): Rect[] {
  if (mode === 'float') return rs.map((r) => ({ x: r.x, y: r.y }));
  return rs.map((r) => {
    const bottom = r.y + r.h;
    const base = lines.reduce(
      (best, y) => (Math.abs(y - bottom) < Math.abs(best - bottom) ? y : best),
      lines[0] ?? 0,
    );
    return { x: r.x, y: base - bottom };
  });
}

/** 逐行增量的偏差（每页第一行取绝对 y），单位 pt */
function lineDeltas(ours: readonly number[], theirs: readonly number[]): number[] {
  return theirs.map((y, k) => {
    const mine = ours[k];
    if (mine === undefined) return Number.POSITIVE_INFINITY;
    const prevMine = ours[k - 1];
    const prevTheirs = theirs[k - 1];
    if (k === 0 || prevMine === undefined || prevTheirs === undefined) return Math.abs(mine - y);
    return Math.abs(mine - prevMine - (y - prevTheirs));
  });
}

describe.each([
  { name: 'spike-image-01', mode: 'inline' as const },
  { name: 'spike-image-02', mode: 'float' as const },
])('$name 与 Word 真值', ({ name, mode }) => {
  const { doc, truth, diagnostics } = layout(name);

  it('页数与 Word 一致，且解析零诊断', () => {
    expect(doc.pages.length).toBe(truth.pageCount);
    expect(diagnostics).toBe(0);
  });

  it('逐行：行高与 Word 一致（比增量，见文件头 ①）', () => {
    doc.pages.forEach((page, i) => {
      const theirs = truth.pages[i]?.lines.map((l) => l.y) ?? [];
      const ours = baselines(page);
      expect(ours.length, `第 ${i + 1} 页行数`).toBe(theirs.length);
      for (const [k, d] of lineDeltas(ours, theirs).entries()) {
        expect(d, `第 ${i + 1} 页第 ${k + 1} 行`).toBeLessThanOrEqual(TOLERANCE_PT);
      }
    });
  });

  it('逐图：位置与 Word 一致', () => {
    doc.pages.forEach((page, i) => {
      const lines = truth.pages[i]?.lines.map((l) => l.y) ?? [];
      const theirs = comparable(
        [...(truth.pages[i]?.images ?? [])]
          .map((im) => ({ x: im.x, y: im.y, h: im.h }))
          .sort((a, b) => a.y - b.y || a.x - b.x),
        lines,
        mode,
      );
      const ours = comparable(rects(page), baselines(page), mode);
      expect(ours.length, `第 ${i + 1} 页图数`).toBe(theirs.length);
      theirs.forEach((t, k) => {
        const m = ours[k] as Rect;
        expect(Math.abs(m.x - t.x), `第 ${i + 1} 页第 ${k + 1} 张图的 x`).toBeLessThanOrEqual(TOLERANCE_PT);
        expect(Math.abs(m.y - t.y), `第 ${i + 1} 页第 ${k + 1} 张图的 y`).toBeLessThanOrEqual(TOLERANCE_PT);
      });
    });
  });
});
