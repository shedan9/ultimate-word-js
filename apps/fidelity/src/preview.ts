#!/usr/bin/env node
/**
 * 把 fixture 画出来看 —— 引擎的第一个「用眼睛验收」的出口。
 *
 *   node src/preview.ts                 # 全部 fixture
 *   node src/preview.ts gongwen-01      # 指定
 *   node src/preview.ts gongwen-01 --truth --debug
 *
 * 产物落在 `apps/fidelity/out/<name>.html`（**不入库**，与 PDF 同理）。
 *
 * `--truth` 会在每一页上叠一层真值：Word 画每一行时的基线用红线标出来，我们自己的
 * 基线是蓝线。两条线重合就是对的，看得见的错位就是 L3 出了问题 —— 这比读一串
 * 「差 0.06pt」的数字快得多，也是「截图回归」这条辅助手段的第一步。
 *
 * 这个脚本**不需要 Word**：docx 与 truth.json 都入库了，度量走随库的度量包，
 * 所以它跨平台。重新造样本才要 Windows。
 */
import { mkdir, readdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDiagnosticSink } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import { layoutDocumentWithFields } from '@uw/layout';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import type { RElement, RenderOptions } from '@uw/render-dom';
import { buildDocument, el, fmt, imageHrefResolver, serialize } from '@uw/render-dom';
import type { WordTruth } from './truth-types.ts';

const APP_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const FIXTURES = path.join(APP_ROOT, 'fixtures');
const OUT = path.join(APP_ROOT, 'out');

interface Args {
  names: string[];
  truth: boolean;
  debug: boolean;
  zoom: number;
}

function parseArgs(argv: readonly string[]): Args {
  const names: string[] = [];
  let truth = false;
  let debug = false;
  let zoom = 1;
  for (const a of argv) {
    if (a === '--truth') truth = true;
    else if (a === '--debug') debug = true;
    else if (a.startsWith('--zoom=')) zoom = Number(a.slice('--zoom='.length)) || 1;
    else if (!a.startsWith('--')) names.push(a.replace(/\.docx$/, ''));
  }
  return { names, truth, debug, zoom };
}

async function allFixtures(): Promise<string[]> {
  const files = await readdir(FIXTURES);
  return files.filter((f) => f.endsWith('.docx')).map((f) => f.slice(0, -'.docx'.length));
}

/**
 * 真值的一层覆盖：每一行的基线画一条红线，横贯整页。
 *
 * 直接往页 `<svg>` 里塞元素，不另建一层 —— viewBox 的单位是 pt，真值的单位也是 pt，
 * 原点还都是纸的左上角，所以真值里的 `y` 拿来就能用，一次换算都不需要。
 * 这不是巧合，是 `paint.ts` 选 pt 做 viewBox 单位的原因之一。
 */
function overlayTruth(root: RElement, truth: WordTruth): void {
  root.children.forEach((svg, i) => {
    const page = truth.pages[i];
    if (page === undefined) return;
    const width = Number((svg.attrs.viewBox ?? '0 0 0 0').split(' ')[2]);
    const lines = page.lines.map((l) =>
      el('line', {
        x1: '0',
        x2: fmt(width),
        y1: fmt(l.y),
        y2: fmt(l.y),
        stroke: '#d1242f',
        'stroke-width': '0.2',
        'stroke-dasharray': '4 3',
      }),
    );
    svg.children.push(el('g', { class: 'uw-truth' }, lines));
  });
}

/**
 * 一页上我们画出的所有基线（pt，纸左上角原点），去重后按 y 排好。
 *
 * 必须**沿途累加每一层 `<g>` 的 translate**，不能只认版心那一层：页眉页脚是与版心平级的
 * 另外两个框，表格与嵌套行也各带一层偏移。只按版心的原点算，页眉的基线会整体偏出
 * 一个上边距，diff 里表现为「每一行都错」而其实排得没问题。
 */
function baselinesOf(svg: RElement): number[] {
  const ys = new Set<number>();
  const visit = (n: RElement, dy: number): void => {
    // `scale(...)`（`w:w` 横向缩放）也走 transform，匹配不上就当没有偏移
    const m = /translate\(([-\d.]+) ([-\d.]+)\)/.exec(n.attrs.transform ?? '');
    const y0 = dy + (m === null ? 0 : Number(m[2]));
    if (n.tag === 'text' && n.text !== undefined && n.text.trim() !== '') {
      ys.add(Number((y0 + Number(n.attrs.y)).toFixed(2)));
    }
    for (const c of n.children) visit(c, y0);
  };
  visit(svg, 0);
  return [...ys].sort((a, b) => a - b);
}

/** 我们自己画出来的每一行基线（蓝），与红线一起看 */
function overlayOurs(root: RElement): void {
  for (const svg of root.children) {
    const width = Number((svg.attrs.viewBox ?? '0 0 0 0').split(' ')[2]);
    const lines = baselinesOf(svg).map((y) =>
      el('line', {
        x1: '0',
        x2: fmt(width),
        y1: fmt(y),
        y2: fmt(y),
        stroke: '#1f6feb',
        'stroke-width': '0.2',
      }),
    );
    svg.children.push(el('g', { class: 'uw-ours' }, lines));
  }
}

/**
 * 每一页「我们的基线 vs 真值基线」的最大偏差，顺手在命令行里报一句。
 *
 * **这不是断言，别拿它当保真度指标**：配对是按行序号硬配的，只要有一行断得与 Word 不同，
 * 后面每一行都会错位一整行的高度，报出来的就是十几 pt 的假差值。真正的判据在
 * `packages/layout/src/fixture.test.ts`（L2 / L3）与各个 spike 脚本里。
 * 这里的数只用来回答一个问题：「刚才那一改，有没有把某一页整体挪歪」。
 */
function worstDelta(root: RElement, truth: WordTruth): { dy: number; lines: number } {
  let dy = 0;
  let count = 0;
  root.children.forEach((svg, i) => {
    const page = truth.pages[i];
    if (page === undefined) return;
    baselinesOf(svg).forEach((y, k) => {
      const t = page.lines[k];
      if (t === undefined) return;
      count++;
      dy = Math.max(dy, Math.abs(y - t.y));
    });
  });
  return { dy, lines: count };
}

const STYLE = `
  body { margin: 0; padding: 24px; background: #6e7781; font-family: system-ui, sans-serif; }
  .uw-doc { display: flex; flex-direction: column; align-items: center; gap: 24px; }
  .uw-page { box-shadow: 0 2px 12px rgba(0,0,0,.35); background: #fff; }
  .uw-page-filler { opacity: .6; }
  figcaption { color: #f6f8fa; font-size: 13px; margin: 0 0 12px; text-align: center; }
`;

async function preview(name: string, args: Args): Promise<void> {
  const sink = createDiagnosticSink();
  const bytes = new Uint8Array(await readFile(path.join(FIXTURES, `${name}.docx`)));
  const doc = loadDocument(OpcPackage.open(bytes), sink);

  const registry = new FontRegistry();
  for (const pack of loadBundledPacks()) registry.registerMetrics(pack);
  const measurer = createTextMeasurer(registry, {
    candidates: (family) => fontNameCandidates(doc.fonts, family),
    diagnostics: sink,
  });

  // 走带域求值的那条：PAGE / NUMPAGES 要是显示成文件里存的旧值，
  // 叠真值时会以为是分页错了
  const { layout: paged } = layoutDocumentWithFields(doc.resolved, doc.fields, {
    measurer,
    settings: doc.cascade.settings,
    headerFooters: doc.headerFooters,
    diagnostics: sink,
  });
  // 图片走 data URI：产物是一份能直接双击打开的 HTML，外部文件一个都不带
  const opts: RenderOptions = {
    zoom: args.zoom,
    debug: args.debug,
    imageHref: imageHrefResolver(doc.images),
  };
  const root = buildDocument(paged, opts);

  let note = `${paged.pages.length} 页`;
  if (args.truth) {
    const raw = await readFile(path.join(FIXTURES, `${name}.truth.json`), 'utf8').catch(() => undefined);
    if (raw === undefined) {
      note += ' · 没有 truth.json，跳过叠加';
    } else {
      const truth = JSON.parse(raw) as WordTruth;
      overlayTruth(root, truth);
      overlayOurs(root);
      const { dy, lines } = worstDelta(root, truth);
      note += ` · 真值 ${truth.pageCount} 页 · ${lines} 行基线最大差 ${dy.toFixed(3)}pt`;
    }
  }
  const diags = sink.list();
  if (diags.length > 0) note += ` · ${diags.length} 条诊断`;

  const html = `<!doctype html><html lang="zh"><head><meta charset="utf-8">
<title>${name} · ultimate-word 预览</title><style>${STYLE}</style></head>
<body><figcaption>${name} —— ${note}</figcaption>${serialize(root)}</body></html>`;

  await mkdir(OUT, { recursive: true });
  const file = path.join(OUT, `${name}.html`);
  await writeFile(file, html, 'utf8');
  console.log(`${name.padEnd(20)} ${note}`);
  console.log(`${''.padEnd(20)} → ${path.relative(process.cwd(), file)}`);
  for (const d of diags) console.log(`${''.padEnd(20)}   [${d.code}] ${d.message}`);
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));
  const names = args.names.length > 0 ? args.names : await allFixtures();
  for (const name of names) await preview(name, args);
}

await main();
