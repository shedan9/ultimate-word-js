/**
 * 调试台：把一份 docx 拖进来，看引擎把它画成什么样。
 *
 * 整条链全在浏览器里跑，一个后端调用都没有 —— 而且从 2026-09-13 起它**只认门面**：
 * `UltimateWord.load(bytes)` → `doc.mount(stage)`，与 api.md §1 那两行一字不差。
 * 它原来是把 `OpcPackage.open` → `loadDocument` → `layoutDocumentWithFields` → `mountView`
 * 手接起来的；那条链现在住在 `packages/ultimate-word/src/load.ts` 里，调试台再自己接一遍
 * 只会让门面漏了什么都照不出来。
 *
 * 字体度量走随库分发的**度量包**（`@uw/fonts/packs`，门面在模块加载时就注册好了），
 * 不是浏览器的 `measureText` —— 所以本机装没装仿宋、黑体**不影响排版**，只影响字形好不好看。
 * 这正是自研布局引擎买到的东西，也是这个调试台最值得盯着看的一点：
 * 换一台没有中文字体的机器打开，断行点与基线一个都不会动。
 */
import { twipsToPt } from '@uw/core';
import type { DecorationHandle, DocRange, UwDocument, UwView } from 'ultimate-word';
import { UltimateWord } from 'ultimate-word';

const app = document.querySelector<HTMLElement>('#app');
if (app === null) throw new Error('#app 不在页面上');

app.innerHTML = `
  <header>
    <h1>ultimate-word 调试台</h1>
    <label class="file">选择 docx<input type="file" accept=".docx" hidden></label>
    <label>缩放 <input type="range" min="50" max="300" step="10" value="100"><button type="button" class="fit">适应宽度</button></label>
    <label><input type="checkbox" class="debug"> 画版心与行盒</label>
    <label><input type="checkbox" class="text-layer" checked> 原生选区</label>
    <label><input type="checkbox" class="edit-mode"> 编辑</label>
    <label><input type="checkbox" class="virtualize" checked> 按需绘制</label>
    <label class="find">查找 <input type="search" placeholder="回车到下一处" disabled><span class="hits"></span></label>
    <span class="status">把一份 .docx 拖进来</span>
  </header>
  <div class="stage"></div>
`;

const stage = app.querySelector<HTMLElement>('.stage') as HTMLElement;
const status = app.querySelector<HTMLElement>('.status') as HTMLElement;
const zoomInput = app.querySelector<HTMLInputElement>('input[type=range]') as HTMLInputElement;
const fitButton = app.querySelector<HTMLButtonElement>('.fit') as HTMLButtonElement;
const debugInput = app.querySelector<HTMLInputElement>('.debug') as HTMLInputElement;
const textLayerInput = app.querySelector<HTMLInputElement>('.text-layer') as HTMLInputElement;
const editInput = app.querySelector<HTMLInputElement>('.edit-mode') as HTMLInputElement;
const virtualizeInput = app.querySelector<HTMLInputElement>('.virtualize') as HTMLInputElement;
const fileInput = app.querySelector<HTMLInputElement>('input[type=file]') as HTMLInputElement;
const findInput = app.querySelector<HTMLInputElement>('input[type=search]') as HTMLInputElement;
const hitsLabel = app.querySelector<HTMLElement>('.hits') as HTMLElement;

let doc: UwDocument | undefined;
let view: UwView | undefined;

/**
 * 三个开关（调试框 / 文字层 / 虚拟化）改的是视图的**构造**选项，门面没有 `update()`
 * （重排是 Phase 7 的事），所以是摘掉再挂一个 —— 缩放不走这条，它只改尺寸（架构 §4.1）。
 */
function remount(): void {
  if (doc === undefined) return;
  view?.dispose();
  view = doc.mount(stage, {
    zoom: Number(zoomInput.value) / 100,
    debug: debugInput.checked,
    textLayer: textLayerInput.checked,
    virtualize: virtualizeInput.checked,
    mode: editInput.checked ? 'edit' : 'preview',
  });
  search.rerun();
}

/**
 * 查找 = `doc.find`（模型）→ `view.decorate`（视图）→ `view.scrollTo`（视图）三步，
 * api.md §15 的第一条配方。装饰在缩放后自己跟着走，重挂视图后按同一串字重搜一遍。
 */
const search = (() => {
  let ranges: DocRange[] = [];
  let marks: DecorationHandle[] = [];
  let cursor = -1;
  const clear = () => {
    for (const m of marks) m.dispose();
    marks = [];
    ranges = [];
    cursor = -1;
    hitsLabel.textContent = '';
  };
  const run = (query: string) => {
    clear();
    if (doc === undefined || view === undefined || query.length === 0) return;
    ranges = doc.find(query, { limit: 500 });
    marks = ranges.map((r) => view?.decorate(r, { className: 'hit' }) as DecorationHandle);
    hitsLabel.textContent = ranges.length === 0 ? '无' : `${ranges.length} 处`;
  };
  const go = (step: number) => {
    if (ranges.length === 0 || view === undefined) return;
    marks[cursor]?.dispose();
    if (cursor >= 0) marks[cursor] = view.decorate(ranges[cursor] as DocRange, { className: 'hit' });
    cursor = (cursor + step + ranges.length) % ranges.length;
    marks[cursor]?.dispose();
    marks[cursor] = view.decorate(ranges[cursor] as DocRange, { className: 'hit hit-current' });
    view.scrollTo(ranges[cursor] as DocRange, { align: 'center', behavior: 'smooth' });
    hitsLabel.textContent = `${cursor + 1} / ${ranges.length}`;
  };
  return {
    rerun: () => run(findInput.value),
    clear,
    go,
  };
})();

async function openFile(file: File): Promise<void> {
  status.textContent = `正在读 ${file.name}…`;
  try {
    const t0 = performance.now();
    doc = await UltimateWord.load(file);
    const ms = performance.now() - t0;
    findInput.disabled = false;

    const first = doc.layout.pages[0]?.geometry;
    const size =
      first === undefined
        ? ''
        : ` · ${twipsToPt(first.width).toFixed(0)}×${twipsToPt(first.height).toFixed(0)}pt`;
    const diags = doc.diagnostics;
    status.textContent =
      `${file.name} · ${doc.pageCount} 页${size} · 解析 + 排版 ${ms.toFixed(1)}ms` +
      (diags.length === 0 ? ' · 无诊断' : ` · ${diags.length} 条诊断：${diags[0]?.message ?? ''}`);
    remount();
  } catch (err) {
    // 结构性错误（不是 zip、缺 document.xml）会抛，内容问题只记诊断 —— 原则 1.5
    status.textContent = `打不开 ${file.name}：${err instanceof Error ? err.message : String(err)}`;
  }
}

fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  if (file !== undefined) void openFile(file);
});
zoomInput.addEventListener('input', () => view?.setZoom(Number(zoomInput.value) / 100));
fitButton.addEventListener('click', () => {
  if (view === undefined) return;
  view.setZoom('fit-width');
  zoomInput.value = String(Math.round(view.zoom * 100));
});
debugInput.addEventListener('change', remount);
textLayerInput.addEventListener('change', remount);
virtualizeInput.addEventListener('change', remount);
editInput.addEventListener('change', remount);
findInput.addEventListener('input', search.rerun);
findInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    search.go(e.shiftKey ? -1 : 1);
  }
});

document.addEventListener('dragover', (e) => e.preventDefault());
document.addEventListener('drop', (e) => {
  e.preventDefault();
  const file = e.dataTransfer?.files?.[0];
  if (file !== undefined) void openFile(file);
});
