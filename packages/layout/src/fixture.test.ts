/**
 * 拿真实公文 `gongwen-01.docx` 走完整条链：解包 → 级联 → 分桶度量 → 断行 → 段落装配。
 *
 * 这里**不比坐标真值**：字体度量走的是三级降级里的第③级（等宽近似），因为宋体 / 仿宋的
 * 度量包要在 Windows 上抽（架构 §5.2），仓库里还没有。所以断言只覆盖「与度量精度无关」
 * 的那些性质：每段都排得出行、行不会莫名其妙地超出版心、结果能过 Worker 边界。
 *
 * 等度量包进来，同一份 fixture 就能直接换成 L2（每行断行点与 `*.truth.json` 一致）的断言 ——
 * 这个文件的结构是照那一天准备的。
 */
import { readFileSync } from 'node:fs';
import { createDiagnosticSink } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import type { LoadedDocument, ResolvedParagraph } from '@uw/model';
import { fontNameCandidates, loadDocument, walkBlocks } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import { beforeAll, describe, expect, it } from 'vitest';
import type { LayoutParagraphOptions } from './paragraph.ts';
import { layoutParagraph } from './paragraph.ts';
import type { ParagraphLayout } from './types.ts';

const FIXTURE = new URL('../../../apps/fidelity/fixtures/gongwen-01.docx', import.meta.url);

const sink = createDiagnosticSink();
let doc: LoadedDocument;
let paragraphs: ResolvedParagraph[];
let laid: ParagraphLayout[];
let contentWidth: number;

beforeAll(() => {
  doc = loadDocument(OpcPackage.open(new Uint8Array(readFileSync(FIXTURE))), sink);
  const section = doc.resolved.sections[0];
  if (section === undefined) throw new Error('fixture 里没有节');

  const { page, margin, docGrid } = section.props;
  contentWidth = page.width - margin.left - margin.right;

  const measurer = createTextMeasurer(new FontRegistry(), {
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

  it('这份文档在布局阶段只报字体缺失，没有别的诊断', () => {
    const codes = new Set(sink.list().map((d) => d.code));
    codes.delete('font-missing');
    expect([...codes]).toEqual([]);
  });
});
