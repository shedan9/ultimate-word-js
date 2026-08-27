/**
 * 表格几何与条件格式的真值回归：`spike-table-01/02` 两份样本，**逐格**与 Word 对。
 *
 * 这是表格层的第一份真值 —— 在它之前列宽、格内边距、行高、`w:vAlign`、跨列的可用宽
 * 与条件格式的层序**全部只有规范做依据**。规则怎么标定出来的见 `apps/fidelity` 的
 * `spike:table`（`TableRules` 的 3 × 2 种组合逐段跑，唯一满分）。
 *
 * 与 `image-fixture.test.ts` 同样两处：
 * ① **y 比的是逐段增量**，不是累加的绝对值 —— Word 自己的行位置带 ±0.13pt 抖动，
 *    32 段叠下来越过判据，而那是它内部取整的锅。格线规则决定的恰好就是增量：
 *    跨一条格线时那一步比行距多出多少；
 * ② 每页第一段仍按绝对 y 比。
 *
 * `spike-table-02` 那份还逐格比**字号** —— 每个条件格式在样式里设了一个独一无二的字号，
 * 于是「这一格最终几号字」= 「层序里最后一个命中的条件是谁」，直接读得出来。
 *
 * 跨平台：docx 与 truth.json 都入库，度量走随库的度量包，所以 CI 上也跑得了。
 */
import { readFileSync } from 'node:fs';
import { createDiagnosticSink } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { layoutDocumentWithFields } from './fields.ts';
import type { PageLayout } from './page.ts';
import type { Chunk } from './test-flatten.ts';
import { chunksOf } from './test-flatten.ts';

const FIXTURES = new URL('../../../apps/fidelity/fixtures/', import.meta.url);
/** L3 / L4 的判据 */
const TOLERANCE_PT = 0.5;
/** 字号的判据：真值的 size 是从文本矩阵反推的，实测 ±0.05pt 抖动 */
const SIZE_TOLERANCE_PT = 0.2;
/**
 * 允许对不上的段数。**只许往下调**。
 *
 * 现在唯一对不上的是 `spike-table-01` 里 `w:tcMar left="0"` 的那一格：
 * Word 把文字放在格边右边 0.59pt 处，我们放在格边上。同一列里边距 5.4pt 的那几格
 * 只差 0.32pt、20pt 的差 0.24pt —— 三个数凑不出一条规则，没有模型就不硬凑。
 */
const ALLOWED_MISSES = 1;

interface TruthItem {
  x: number;
  y: number;
  w: number;
  size: number;
  text: string;
}
interface Truth {
  pageCount: number;
  pages: { items: TruthItem[] }[];
}

/**
 * 同一基线上、x 首尾相接的片段并成一段。两侧共用同一套并法，切法才对得上 ——
 * pdf.js 按字体子集切，我们按 run 切，逐片段配对是配不上的。
 */
function merge(pieces: readonly Chunk[]): Chunk[] {
  const sorted = [...pieces].filter((p) => p.text.trim() !== '').sort((a, b) => a.y - b.y || a.x - b.x);
  const out: Chunk[] = [];
  for (const p of sorted) {
    const last = out[out.length - 1];
    if (last !== undefined && Math.abs(last.y - p.y) < 0.6 && p.x - last.right < 2) {
      last.text += p.text;
      last.right = p.right;
      continue;
    }
    out.push({ ...p });
  }
  return out;
}

/** 一页的全部文字片段，坐标相对**纸**左上角，并完段 */
function pageChunks(page: PageLayout): Chunk[] {
  return merge(chunksOf(page));
}

function truthChunks(truth: Truth, page: number): Chunk[] {
  return merge(
    (truth.pages[page]?.items ?? []).map((i) => ({
      x: i.x,
      y: i.y,
      right: i.x + i.w,
      size: i.size,
      text: i.text,
    })),
  );
}

function layout(name: string): { pages: PageLayout[]; truth: Truth } {
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
    pages: layoutDocumentWithFields(loaded.resolved, loaded.fields, {
      measurer,
      settings: loaded.cascade.settings,
      headerFooters: loaded.headerFooters,
    }).layout.pages,
    truth: JSON.parse(readFileSync(new URL(`${name}.truth.json`, FIXTURES), 'utf8')) as Truth,
  };
}

describe.each(['spike-table-01', 'spike-table-02'])('%s 逐格与 Word 对', (name) => {
  const { pages, truth } = layout(name);

  it('页数一致', () => {
    expect(pages).toHaveLength(truth.pageCount);
  });

  it('每一段的文字一致', () => {
    for (let p = 0; p < truth.pageCount; p++) {
      const want = truthChunks(truth, p).map((c) => c.text);
      const page = pages[p];
      expect(page).toBeDefined();
      expect(pageChunks(page as PageLayout).map((c) => c.text)).toEqual(want);
    }
  });

  it('每一段的 x / 增量 y / 字号都在判据内', () => {
    const misses: string[] = [];
    for (let p = 0; p < truth.pageCount; p++) {
      const want = truthChunks(truth, p);
      const got = pageChunks(pages[p] as PageLayout);
      for (let i = 0; i < want.length; i++) {
        const w = want[i] as Chunk;
        const m = got[i] as Chunk;
        const prevW = want[i - 1];
        const prevM = got[i - 1];
        const dy =
          prevW === undefined || prevM === undefined
            ? Math.abs(m.y - w.y)
            : Math.abs(m.y - prevM.y - (w.y - prevW.y));
        const dx = Math.abs(m.x - w.x);
        const ds = Math.abs(m.size - w.size);
        if (dx > TOLERANCE_PT || dy > TOLERANCE_PT || ds > SIZE_TOLERANCE_PT) {
          misses.push(`「${w.text}」Δx=${dx.toFixed(2)} Δy=${dy.toFixed(2)} Δ字号=${ds.toFixed(2)}`);
        }
      }
    }
    expect(misses.length, misses.join('; ')).toBeLessThanOrEqual(ALLOWED_MISSES);
  });
});

describe('spike-table-01 的几条几何结论各留一条断言', () => {
  const { pages, truth } = layout('spike-table-01');
  const got = pageChunks(pages[0] as PageLayout);
  const want = truthChunks(truth, 0);
  const find = (list: readonly Chunk[], text: string): Chunk => {
    const c = list.find((x) => x.text === text);
    if (c === undefined) throw new Error(`没有「${text}」这一段`);
    return c;
  };

  it('6pt 粗边框不吃格内可用宽 —— 那一格照样一行 9 个字', () => {
    // 吃了的话可用宽从 109.2pt 掉到 97.2pt，一行只剩 8 个字，这一段的文字就不一样了
    expect(find(got, '戊一二三四五六七八').text).toBe(find(want, '戊一二三四五六七八').text);
  });

  it('跨两列的格子按合并后的宽度断行 —— 一行 16 个字', () => {
    expect(find(got, '丙一二三四五六七八九十子丑寅卯辰')).toBeDefined();
  });

  it('w:vAlign 的三种摆法：free 空间的 0 / 一半 / 全部', () => {
    const top = find(got, '丁一').y;
    const center = find(got, '丁二').y;
    const bottom = find(got, '丁三').y;
    const wTop = find(want, '丁一').y;
    expect(Math.abs(center - top - (find(want, '丁二').y - wTop))).toBeLessThan(TOLERANCE_PT);
    expect(Math.abs(bottom - top - (find(want, '丁三').y - wTop))).toBeLessThan(TOLERANCE_PT);
    // 居中正好是靠底那一半：行高 60pt、内容 15.6pt，富余 44.4pt
    expect(Math.abs((center - top) * 2 - (bottom - top))).toBeLessThan(TOLERANCE_PT);
  });

  it('6pt 的格线在纵向占满一整条线宽', () => {
    // 第 5 行第一格四条边都是 6pt：它上边那条格线把整行往下推 6pt（不是 3pt，也不是 0）
    const before = find(got, '巳午未申').y; // 第 3 行的第二行文字
    const row4 = find(got, '丁一').y;
    const row5 = find(got, '戊一二三四五六七八').y;
    const wBefore = find(want, '巳午未申').y;
    expect(Math.abs(row4 - before - (find(want, '丁一').y - wBefore))).toBeLessThan(TOLERANCE_PT);
    // 第 4 行 w:trHeight = 60pt，加上 6pt 的格线正好 66pt
    expect(Math.abs(row5 - row4 - 66)).toBeLessThan(TOLERANCE_PT);
  });
});
