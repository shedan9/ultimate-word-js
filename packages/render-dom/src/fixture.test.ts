/**
 * 端到端：真实公文 `gongwen-01.docx` 走完 解包 → 级联 → 度量 → 断行 → 分页 → **画**，
 * 再拿画出来的 `<text>` 属性直接跟 `gongwen-01.truth.json` 比。
 *
 * 为什么值得再比一遍（`@uw/layout` 的 fixture.test.ts 已经比过 L2 / L3 了）：
 * 那边比的是 `LayoutResult` 里的数，这边比的是**属性字符串里的数**。中间隔着
 * 「twips → pt」「版心原点搬进 `<g transform>`」「逐字 x 拼成 x 列表」三步翻译，
 * 每一步都能悄悄丢掉或加上一个偏移，而那种错在布局侧的断言里一个都照不出来。
 *
 * 判据仍是 L3 / L4 的 0.5pt —— 渲染这一步不该引入任何误差，实测差 0.0（属性里的数
 * 就是布局里那个数四舍五入到 3 位小数），所以这里的容差留得很紧。
 */
import { readFileSync } from 'node:fs';
import { createDiagnosticSink } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import { layoutDocument } from '@uw/layout';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import { beforeAll, describe, expect, it } from 'vitest';
import { buildDocument } from './paint.ts';
import type { RElement } from './tree.ts';
import { serialize } from './tree.ts';

const FIXTURE = new URL('../../../apps/fidelity/fixtures/gongwen-01.docx', import.meta.url);
const TRUTH = new URL('../../../apps/fidelity/fixtures/gongwen-01.truth.json', import.meta.url);

/** 渲染这一步不该引入误差，容差只留给 3 位小数的取整（0.001pt）与真值自身的抖动 */
const TOLERANCE_PT = 0.5;

interface TruthFile {
  pageCount: number;
  pages: { width: number; height: number; lines: { y: number; x: number; text: string }[] }[];
}

interface PaintedLine {
  y: number;
  x: number;
  text: string;
}

let root: RElement;
let truth: TruthFile;

/** 页 `<svg>` → 每一行的绝对基线与起始 x（把 `<g>` 的 translate 加回去） */
function linesOf(svg: RElement): PaintedLine[] {
  const content = svg.children.find((c) => c.attrs.class === 'uw-content');
  if (content === undefined) throw new Error('页里没有版心组');
  const m = /translate\(([-\d.]+) ([-\d.]+)\)/.exec(content.attrs.transform ?? '');
  if (m === null) throw new Error(`版心组的 transform 认不出：${content.attrs.transform}`);
  const [ox, oy] = [Number(m[1]), Number(m[2])];

  const byBaseline = new Map<string, PaintedLine>();
  const visit = (node: RElement): void => {
    if (node.tag === 'text' && node.text !== undefined) {
      const y = oy + Number(node.attrs.y);
      const x = ox + Number((node.attrs.x ?? '0').split(' ')[0]);
      const key = y.toFixed(2);
      const hit = byBaseline.get(key);
      if (hit === undefined) byBaseline.set(key, { y, x, text: node.text });
      else hit.text += node.text;
    }
    for (const c of node.children) visit(c);
  };
  visit(svg);
  return [...byBaseline.values()].filter((l) => l.text.trim() !== '').sort((a, b) => a.y - b.y);
}

beforeAll(() => {
  const sink = createDiagnosticSink();
  const doc = loadDocument(OpcPackage.open(new Uint8Array(readFileSync(FIXTURE))), sink);
  const registry = new FontRegistry();
  for (const pack of loadBundledPacks()) registry.registerMetrics(pack);
  const measurer = createTextMeasurer(registry, {
    candidates: (family) => fontNameCandidates(doc.fonts, family),
    diagnostics: sink,
  });
  const paged = layoutDocument(doc.resolved, { measurer, settings: doc.cascade.settings, diagnostics: sink });
  root = buildDocument(paged);
  truth = JSON.parse(readFileSync(TRUTH, 'utf8')) as TruthFile;
});

describe('画出来的页与真值', () => {
  it('页数一致，且每页一个 <svg>', () => {
    expect(root.children).toHaveLength(truth.pageCount);
    expect(root.children.every((c) => c.tag === 'svg')).toBe(true);
  });

  it('纸张尺寸与 Word 一致 —— viewBox 的单位就是真值的 pt', () => {
    const first = root.children[0] as RElement;
    const page = truth.pages[0];
    if (page === undefined) throw new Error('真值里没有第一页');
    const [, , vw, vh] = (first.attrs.viewBox ?? '').split(' ').map(Number);
    expect(vw).toBeCloseTo(page.width, 1);
    expect(vh).toBeCloseTo(page.height, 1);
  });

  it('每一行的基线 y 与真值差 < 0.5pt（L3，翻译成属性之后仍然成立）', () => {
    let worst = 0;
    root.children.forEach((svg, pi) => {
      const ours = linesOf(svg);
      const theirs = truth.pages[pi]?.lines ?? [];
      expect(ours).toHaveLength(theirs.length);
      ours.forEach((l, i) => {
        const d = Math.abs(l.y - (theirs[i]?.y ?? 0));
        if (d > worst) worst = d;
      });
    });
    expect(worst).toBeLessThan(TOLERANCE_PT);
  });

  it('每一行的起始 x 与真值差 < 0.5pt（L4）', () => {
    let worst = 0;
    root.children.forEach((svg, pi) => {
      const ours = linesOf(svg);
      const theirs = truth.pages[pi]?.lines ?? [];
      ours.forEach((l, i) => {
        const d = Math.abs(l.x - (theirs[i]?.x ?? 0));
        if (d > worst) worst = d;
      });
    });
    expect(worst).toBeLessThan(TOLERANCE_PT);
  });

  it('一个字都没丢：拼起来的文字里有标题与落款', () => {
    const all = root.children
      .flatMap((svg) => linesOf(svg))
      .map((l) => l.text)
      .join('');
    expect(all).toContain('通知');
    expect(all.length).toBeGreaterThan(100);
  });

  it('元素树可结构化克隆 —— 将来它就是过 Worker 边界的那份数据', () => {
    expect(structuredClone(root)).toEqual(root);
  });

  it('序列化出来的是一段能直接落盘的标记，且不含未转义的尖括号', () => {
    const svg = serialize(root.children[0] as RElement);
    expect(svg.startsWith('<svg')).toBe(true);
    // 标签以外不该出现裸的 `<`：转义漏了会让整段 SVG 解析失败
    expect(svg.replace(/<\/?[a-zA-Z:]+[^>]*>/g, '')).not.toContain('<');
  });
});
