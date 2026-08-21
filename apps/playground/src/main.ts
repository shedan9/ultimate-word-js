/**
 * 调试台：把一份 docx 拖进来，看引擎把它画成什么样。
 *
 * 整条链全在浏览器里跑，一个后端调用都没有：
 * `OpcPackage.open` → `loadDocument` → `layoutDocument` → `mount`。
 *
 * 字体度量走随库分发的**度量包**（`packages/fonts/packs/*.json`），不是浏览器的
 * `measureText` —— 所以本机装没装仿宋、黑体**不影响排版**，只影响字形好不好看。
 * 这正是自研布局引擎买到的东西，也是这个调试台最值得盯着看的一点：
 * 换一台没有中文字体的机器打开，断行点与基线一个都不会动。
 *
 * `import.meta.glob` 是 Vite 专属语法，只出现在这个 app 里 —— `@uw/fonts` 的主入口
 * 刻意不依赖任何打包器（也不依赖 fontkit），度量包在浏览器侧怎么送进去是**调用方**
 * 的事。真产品里这一步多半是 `fetch` 一个合并后的包，不是 17 个 JSON。
 */
import { createDiagnosticSink, twipsToPt } from '@uw/core';
import type { MetricsPack } from '@uw/fonts';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { layoutDocument } from '@uw/layout';
import { fontNameCandidates, loadDocument } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import { mount } from '@uw/render-dom/dom';

const packs = import.meta.glob<MetricsPack>('../../../packages/fonts/packs/*.json', {
  eager: true,
  import: 'default',
});

/** 注册表建一次就够，度量包与文档无关 */
const registry = new FontRegistry();
for (const [file, pack] of Object.entries(packs)) {
  if (file.endsWith('index.json')) continue;
  registry.registerMetrics(pack);
}

const app = document.querySelector<HTMLElement>('#app');
if (app === null) throw new Error('#app 不在页面上');

app.innerHTML = `
  <header>
    <h1>ultimate-word 调试台</h1>
    <label class="file">选择 docx<input type="file" accept=".docx" hidden></label>
    <label>缩放 <input type="range" min="50" max="200" step="10" value="100"></label>
    <label><input type="checkbox" class="debug"> 画版心与行盒</label>
    <span class="status">把一份 .docx 拖进来</span>
  </header>
  <div class="stage"></div>
`;

const stage = app.querySelector<HTMLElement>('.stage') as HTMLElement;
const status = app.querySelector<HTMLElement>('.status') as HTMLElement;
const zoomInput = app.querySelector<HTMLInputElement>('input[type=range]') as HTMLInputElement;
const debugInput = app.querySelector<HTMLInputElement>('.debug') as HTMLInputElement;
const fileInput = app.querySelector<HTMLInputElement>('input[type=file]') as HTMLInputElement;

/** 当前文档的布局结果。缩放与调试开关只重画，**不重排** —— 架构 §4.1 */
let current: ReturnType<typeof layoutDocument> | undefined;

function draw(): void {
  if (current === undefined) return;
  mount(stage, current, { zoom: Number(zoomInput.value) / 100, debug: debugInput.checked });
}

function open(bytes: Uint8Array, name: string): void {
  const t0 = performance.now();
  const sink = createDiagnosticSink();
  const doc = loadDocument(OpcPackage.open(bytes), sink);
  const measurer = createTextMeasurer(registry, {
    candidates: (family) => fontNameCandidates(doc.fonts, family),
    diagnostics: sink,
  });
  current = layoutDocument(doc.resolved, {
    measurer,
    settings: doc.cascade.settings,
    diagnostics: sink,
  });
  const ms = performance.now() - t0;

  const first = current.pages[0]?.geometry;
  const size =
    first === undefined
      ? ''
      : ` · ${twipsToPt(first.width).toFixed(0)}×${twipsToPt(first.height).toFixed(0)}pt`;
  const diags = sink.list();
  status.textContent =
    `${name} · ${current.pages.length} 页${size} · 解析 + 排版 ${ms.toFixed(1)}ms` +
    (diags.length === 0 ? ' · 无诊断' : ` · ${diags.length} 条诊断：${diags[0]?.message ?? ''}`);
  draw();
}

async function openFile(file: File): Promise<void> {
  status.textContent = `正在读 ${file.name}…`;
  try {
    open(new Uint8Array(await file.arrayBuffer()), file.name);
  } catch (err) {
    // 结构性错误（不是 zip、缺 document.xml）会抛，内容问题只记诊断 —— 原则 1.5
    status.textContent = `打不开 ${file.name}：${err instanceof Error ? err.message : String(err)}`;
  }
}

fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  if (file !== undefined) void openFile(file);
});
zoomInput.addEventListener('input', draw);
debugInput.addEventListener('change', draw);

document.addEventListener('dragover', (e) => e.preventDefault());
document.addEventListener('drop', (e) => {
  e.preventDefault();
  const file = e.dataTransfer?.files?.[0];
  if (file !== undefined) void openFile(file);
});
