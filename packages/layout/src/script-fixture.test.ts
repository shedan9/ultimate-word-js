/**
 * 脚本分桶与行盒合成的真值回归：`spike-script-01` 一份样本 11 页，**逐页逐行**与 Word 对基线。
 *
 * 这份样本回答的两问（怎么标定出来的见 `apps/fidelity` 的 `spike:script`，规则表在
 * `line-height.ts` 的 `SCRIPT_RULES`）：
 *
 * 1. 行高走东亚规则（×1.3）还是拉丁规则（GDI 外部行距），看的是**实际画字的那款字体**，
 *    逐段判 —— 不是这一行有没有东亚字符（旧实现），也不是 `w:eastAsia` 槽或 `w:hint`
 * 2. 同一行里几款字体，**各自的行盒逐项取 max**（上取最高、下取最深）——
 *    不是「取各自行高的最大值」
 *
 * 每页四段同格式的短段连排，所以「相邻基线差」就是行高；每页换一种「ascii 槽 × eastAsia 槽」
 * 的配法。P1–P7 正文是**纯 ASCII**（第 1 问），P8–P11 带汉字（第 2 问 + 对照）。
 *
 * 比的是**绝对基线 y**：这份样本一页只有四行，Word 自己那 ±0.1pt 的行距抖动累不到判据上
 * （实测最坏 0.368pt），不必像 `spike:image` 那样退到逐行增量。
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
import type { DocumentLayout, PageLayout } from './page.ts';

const FIXTURES = new URL('../../../apps/fidelity/fixtures/', import.meta.url);
const NAME = 'spike-script-01';
/** L3 的判据。四种说法在这份样本上相差 3.5–14pt，0.5pt 分得干干净净 */
const TOLERANCE_PT = 0.5;

interface Truth {
  pageCount: number;
  pages: { lines: { y: number; text: string }[] }[];
}

/**
 * 比文字时把空白全去掉：真值里的「P8 汉 TimesAsciiDengEa」多出来的两个空格是
 * **中西文自动间距**在 PDF 里张开的缝，抽真值那一步按片段间距补的。它是宽度那一维的事，
 * 这份样本量的是基线 y。
 */
const squash = (s: string): string => s.replace(/\s+/gu, '');

function layout(): { doc: DocumentLayout; truth: Truth; diagnostics: number } {
  const sink = createDiagnosticSink();
  const bytes = new Uint8Array(readFileSync(new URL(`${NAME}.docx`, FIXTURES)));
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
    }).layout,
    truth: JSON.parse(readFileSync(new URL(`${NAME}.truth.json`, FIXTURES), 'utf8')) as Truth,
    diagnostics: sink.list().length,
  };
}

/** 一页里每一行的绝对基线（pt）与文字 */
function pageLines(page: PageLayout): { y: number; text: string }[] {
  const out: { y: number; text: string }[] = [];
  for (const b of page.blocks) {
    if (b.kind !== 'paragraph') continue;
    for (const placed of b.lines) {
      const text = placed.line.fragments
        .map((f) => f.text)
        .join('')
        .replace(/\s+$/u, '');
      if (text !== '')
        out.push({ y: (page.geometry.content.y + placed.y + placed.line.baseline) / 20, text });
    }
  }
  return out;
}

describe(`${NAME} · 纯 ASCII 的行走哪套规则 + 几款字体怎么合成行盒`, () => {
  const { doc, truth, diagnostics } = layout();

  it('L0：页数与 Word 一致', () => {
    expect(doc.pages).toHaveLength(truth.pageCount);
  });

  it('每一页的每一行都与 Word 逐字一致', () => {
    const ours = doc.pages.map((p) => pageLines(p).map((l) => squash(l.text)));
    const theirs = truth.pages.map((p) => p.lines.map((l) => squash(l.text)));
    expect(ours).toEqual(theirs);
  });

  it(`L3：每一行的基线 y 与真值差 < ${TOLERANCE_PT}pt`, () => {
    const worst = { page: -1, line: -1, delta: 0, text: '' };
    doc.pages.forEach((p, i) => {
      pageLines(p).forEach((l, k) => {
        const delta = Math.abs(l.y - (truth.pages[i]?.lines[k]?.y ?? l.y));
        if (delta > worst.delta) Object.assign(worst, { page: i, line: k, delta, text: l.text });
      });
    });
    expect(
      worst.delta,
      `最坏的是第 ${worst.page + 1} 页第 ${worst.line + 1} 行「${worst.text}」`,
    ).toBeLessThan(TOLERANCE_PT);
  });

  /**
   * 把「按字体判」这条单独钉一遍：整份样本的行高一起比只能说明「合起来对了」，
   * 而这一条是**改了实现**的那一条 —— 旧实现（按行里有没有东亚字符判）在 P2 上给 41.06pt。
   * 一页四段，所以相邻基线差就是行高。
   */
  it('纯 ASCII 的一行，行高跟着 ascii 槽里那款字体走', () => {
    const heights = doc.pages.map((p) => {
      const ls = pageLines(p);
      return Math.round(((ls[1]?.y ?? 0) - (ls[0]?.y ?? 0)) * 100) / 100;
    });
    // P1/P6 的 ascii 槽是 Times（拉丁规则 41.40），P2/P5 是宋体 / 仿宋（东亚规则 46.80），
    // P3 等线 48.77，P4/P7 微软雅黑 61.77 —— 旧实现这七页全给拉丁值
    expect(heights.slice(0, 7)).toEqual([41.4, 46.8, 48.77, 61.77, 46.8, 41.4, 61.77]);
  });

  /**
   * 合成规则同理单独钉一遍：P9–P11 是等线与宋体 / 仿宋同行，Word 给的行高
   * **比两款字体各自的行高都大**（50.28 > 48.77 > 46.80）——「取行高最大值」说不出这个数。
   */
  it('两款东亚字体同行时，行盒上取最高、下取最深', () => {
    const heights = doc.pages.slice(8).map((p) => {
      const ls = pageLines(p);
      return Math.round(((ls[1]?.y ?? 0) - (ls[0]?.y ?? 0)) * 100) / 100;
    });
    expect(heights).toEqual([50.31, 50.31, 50.31]);
  });

  it('一条诊断都没有', () => {
    expect(diagnostics).toBe(0);
  });
});
