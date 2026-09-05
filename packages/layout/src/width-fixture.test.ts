/**
 * 宽度分桶与中西文间距的真值回归：`spike-width-01` 一份样本 3 页 31 行，
 * **逐行**与 Word 对「片段的字体序列」与「行末 x」。
 *
 * 这份样本回答的三问（怎么标定出来的见 `apps/fidelity` 的 `spike:width`，
 * 规则表与证据在 `items.ts` 的 `WIDTH_RULES`）：
 *
 * 1. 歧义字符（EastAsianWidth = A：`§ ° ± × ÷ ·`…）跟着 `w:hint` 走，**与邻居无关**
 * 2. 空格随东亚邻居（任一侧），**与 `w:hint` 无关**（旧实现要求 hint=eastAsia，是猜的）；
 *    `/` `-` 这类中性字符不随
 * 3. 中西文自动间距是 **1/4 em**（不是开发计划记的 1/8），按**接缝前面**那个字符的字号算，
 *    且靠 hint 才进东亚桶的歧义字符不加
 *
 * 为什么比这两样：真值的 `TruthItem.font` 直接说出 Word 用哪款字体画了这个字
 * （字体一换，PDF 里就换一次 `Tf`、起一个新片段），所以分桶是**读**出来的；
 * 而分桶与间距错了都落在行末 x 上，一个数把两种错都兜住。
 * 唯一读不得字体名的是**空格** —— Word 画它时不换 `Tf`（见 `@uw/fonts` 的
 * `neutralTakesEastAsia`），那几行只能靠行末 x 分辨，好在 0.5 em 与 0.25 em 差 9pt。
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
const NAME = 'spike-width-01';
/** L4 的判据。各条假设互相差 9–36pt，0.5pt 分得干干净净（实测最坏 0.089pt） */
const TOLERANCE_PT = 0.5;

/**
 * 文档里的字体名 → 真值里的 PostScript 名。
 * 写死而不是从 `fontTable.xml` 推：这份样本只用两款字体，而「本地化名 → 磁盘字体 →
 * PostScript 名」那条链本身还没标定，混进来会让读数依赖一个未知量。
 */
const PS_NAME: Record<string, string> = {
  宋体: 'SimSun',
  'Times New Roman': 'TimesNewRomanPSMT',
};

interface TruthItem {
  font: string;
  text: string;
}
interface Truth {
  pageCount: number;
  pages: { items: TruthItem[]; lines: { xEnd: number; text: string; items: number[] }[] }[];
}

const squash = (s: string): string => s.replace(/\s+/gu, '');

/** 相邻重复的字体合并掉 —— 我们按 run 切片段，Word 按 `Tf` 切，粒度本来就不同 */
function dedupe(fonts: readonly string[]): string[] {
  const out: string[] = [];
  for (const f of fonts) {
    if (out[out.length - 1] !== f) out.push(f);
  }
  return out;
}

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

interface Line {
  xEnd: number;
  text: string;
  fonts: string[];
}

function pageLines(page: PageLayout): Line[] {
  const left = page.geometry.content.x;
  const out: Line[] = [];
  for (const b of page.blocks) {
    if (b.kind !== 'paragraph') continue;
    for (const placed of b.lines) {
      const frags = placed.line.fragments.filter((f) => f.text !== '');
      if (frags.length === 0) continue;
      const text = frags
        .map((f) => f.text)
        .join('')
        .replace(/\s+$/u, '');
      if (text === '') continue;
      out.push({
        xEnd: (left + Math.max(...frags.map((f) => f.x + f.width))) / 20,
        text,
        // 空白片段不参与：Word 画空格时不换 Tf，真值里它跟着邻居的字体走
        fonts: dedupe(frags.filter((f) => f.text.trim() !== '').map((f) => PS_NAME[f.font] ?? f.font)),
      });
    }
  }
  return out;
}

function truthLines(truth: Truth, pageIndex: number): Line[] {
  const page = truth.pages[pageIndex];
  if (page === undefined) return [];
  return page.lines.map((l) => ({
    xEnd: l.xEnd,
    text: l.text,
    fonts: dedupe(
      l.items
        .map((i) => page.items[i])
        .filter((it) => it !== undefined && it.text.trim() !== '')
        .map((it) => (it as TruthItem).font),
    ),
  }));
}

describe(`${NAME} · 歧义字符与中性字符进哪个桶 + 中西文自动间距`, () => {
  const { doc, truth, diagnostics } = layout();

  it('L0：页数与 Word 一致', () => {
    expect(doc.pages).toHaveLength(truth.pageCount);
  });

  it('L2：每一页的每一行都与 Word 逐字一致', () => {
    const ours = doc.pages.map((p) => pageLines(p).map((l) => squash(l.text)));
    const theirs = truth.pages.map((_, i) => truthLines(truth, i).map((l) => squash(l.text)));
    expect(ours).toEqual(theirs);
  });

  it('每一行的片段字体序列与 Word 一致 —— 分桶结果是读出来的，不是从宽度反推的', () => {
    const ours = doc.pages.map((p) => pageLines(p).map((l) => l.fonts.join('>')));
    const theirs = truth.pages.map((_, i) => truthLines(truth, i).map((l) => l.fonts.join('>')));
    expect(ours).toEqual(theirs);
  });

  it(`L4：每一行的行末 x 与真值差 < ${TOLERANCE_PT}pt`, () => {
    const worst = { page: -1, line: -1, delta: 0, text: '' };
    doc.pages.forEach((p, i) => {
      const theirs = truthLines(truth, i);
      pageLines(p).forEach((l, k) => {
        const delta = Math.abs(l.xEnd - (theirs[k]?.xEnd ?? l.xEnd));
        if (delta > worst.delta) Object.assign(worst, { page: i, line: k, delta, text: l.text });
      });
    });
    expect(
      worst.delta,
      `最坏的是第 ${worst.page + 1} 页第 ${worst.line + 1} 行「${worst.text}」`,
    ).toBeLessThan(TOLERANCE_PT);
  });

  /**
   * 把「歧义字符跟着 hint」单独钉一遍：Ea4 与 De4 的文字一模一样（`B§B°B±B`），
   * 只有 `w:hint` 不同 —— 一个把 `§ ° ±` 交给宋体（各 36pt），一个交给 Times（18 / 14.4 / 19.76pt）。
   * 「跟邻居走」那条说法在这一对上就被打掉了，所以这两行值得单拎出来。
   */
  it('歧义字符跟着 w:hint 走：同样的 B§B°B±B，两种 hint 给出相反的分桶', () => {
    const all = doc.pages.flatMap((p) => pageLines(p));
    const ea = all.find((l) => l.text.startsWith('Ea4'));
    const de = all.find((l) => l.text.startsWith('De4'));
    expect(ea?.fonts).toEqual([
      'TimesNewRomanPSMT',
      'SimSun',
      'TimesNewRomanPSMT',
      'SimSun',
      'TimesNewRomanPSMT',
      'SimSun',
      'TimesNewRomanPSMT',
      'SimSun',
      'TimesNewRomanPSMT',
    ]);
    expect(de?.fonts).toEqual(['TimesNewRomanPSMT', 'SimSun', 'TimesNewRomanPSMT']);
  });

  /**
   * 自动间距那一条也单独钉：As1（`中B中`）与 As3（同样的形状但关掉 `w:autoSpaceDE`）
   * 的行末 x 差的正好是两个接缝的间距。1/4 与 1/8 在这里差 9pt，
   * 而这一条是**改了实现**的那一条（原来写的是 1/8）。
   */
  it('中西文自动间距 = 1/4 em：开关一关，一行少两个 9pt', () => {
    const all = doc.pages.flatMap((p) => pageLines(p));
    const on = all.find((l) => l.text.startsWith('As1'));
    const off = all.find((l) => l.text.startsWith('As3'));
    // 36pt 字：两个接缝各 9pt。B 与 C 的字形宽度相同（Times 都是 0.667 em）
    expect(Math.round(((on?.xEnd ?? 0) - (off?.xEnd ?? 0)) * 100) / 100).toBe(18);
  });

  it('一条诊断都没有', () => {
    expect(diagnostics).toBe(0);
  });
});
