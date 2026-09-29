<script setup lang="ts">
/**
 * 在线预览：整条路径只走门面包（`UltimateWord.load` → `doc.mount`），与调试台同一套接线，
 * 但只摆给使用者看的那几个开关 —— 版心框、虚拟化这类调试项留在 apps/playground。
 * 门面在 onMounted 里动态 import：VitePress 要先在 Node 里 SSR 一遍，库的 DOM 部分不能进那一趟。
 */
import type {
  DecorationHandle,
  DocRange,
  UltimateWord as UltimateWordApi,
  UwDocument,
  UwView,
} from 'ultimate-word';
import { onBeforeUnmount, onMounted, ref, shallowRef, watch } from 'vue';
import gongwenUrl from '../../../fidelity/fixtures/gongwen-01.docx?url';
import headerUrl from '../../../fidelity/fixtures/spike-header-03.docx?url';
import imageUrl from '../../../fidelity/fixtures/spike-image-01.docx?url';
import tableUrl from '../../../fidelity/fixtures/spike-table-04.docx?url';

interface Sample {
  id: string;
  label: string;
  hint: string;
  url: string;
}

// 样本都是真值流水线的 fixture：每一份都与 Word 导出的 PDF 逐行比过，看到的就是「对上了」的那一版
const samples: Sample[] = [
  { id: 'gongwen-01', label: '公文', hint: '真实公文：标点挤压、悬挂、中西文间距', url: gongwenUrl },
  { id: 'spike-table-04', label: '表格跨页', hint: '一行放不下时从行间切开，续页重复表头', url: tableUrl },
  { id: 'spike-image-01', label: '内嵌图片', hint: '图片盒量化到 1.5pt、坐在基线上', url: imageUrl },
  {
    id: 'spike-header-03',
    label: '页眉页脚',
    hint: '首页 / 奇偶页各一份，页脚里是真的 PAGE 域',
    url: headerUrl,
  },
];

const stage = ref<HTMLElement>();
const api = shallowRef<typeof UltimateWordApi>();
const doc = shallowRef<UwDocument>();
const view = shallowRef<UwView>();

const current = ref<string>(samples[0]?.id ?? '');
const fileName = ref('');
const loading = ref(false);
const error = ref('');
const dragging = ref(false);
const zoom = ref(100);
const editing = ref(false);
const query = ref('');
const hits = ref<{ total: number; at: number }>({ total: 0, at: -1 });
const stats = ref<{ pages: number; size: string; ms: number; diagnostics: string[] }>();

let docEvents: Array<{ dispose(): void }> = [];

async function open(source: string | File, name: string, id: string): Promise<void> {
  if (api.value === undefined) return;
  loading.value = true;
  error.value = '';
  try {
    const t0 = performance.now();
    const next = await api.value.load(source);
    const ms = performance.now() - t0;
    teardown();
    doc.value = next;
    current.value = id;
    fileName.value = name;
    refreshStats(ms);
    docEvents = [
      next.on('layout:done', ({ duration }) => {
        refreshStats(duration);
        runSearch();
      }),
    ];
    mount(true);
  } catch (err) {
    // 结构性错误（不是 zip、缺 document.xml）会抛；内容问题只记诊断，照样画
    error.value = `打不开 ${name}：${err instanceof Error ? err.message : String(err)}`;
  } finally {
    loading.value = false;
  }
}

function refreshStats(ms: number): void {
  const d = doc.value;
  if (d === undefined) return;
  const first = d.layout.pages[0]?.geometry;
  // 布局单位是 twips，1pt = 20 twips；这里只为显示，不经 @uw/core 免得 demo 多一个依赖
  const size =
    first === undefined ? '—' : `${Math.round(first.width / 20)} × ${Math.round(first.height / 20)} pt`;
  stats.value = { pages: d.pageCount, size, ms, diagnostics: d.diagnostics.map((x) => x.message) };
}

/**
 * 视图没有 `update()`：编辑开关是构造选项，改它就摘掉重挂（api.md §3）。
 * 首次挂载按容器宽度适配，但不放大 —— 宽屏上把 A4 撑到 140% 反而看不出版面。
 */
function mount(fit: boolean): void {
  const d = doc.value;
  if (d === undefined || stage.value === undefined) return;
  clearSearch();
  view.value?.dispose();
  const v = d.mount(stage.value, { zoom: zoom.value / 100, mode: editing.value ? 'edit' : 'preview' });
  view.value = v;
  if (fit) {
    v.setZoom('fit-width');
    if (v.zoom > 1) v.setZoom(1);
    zoom.value = Math.round(v.zoom * 100);
  }
  // 目录条目是 `HYPERLINK \l "_Toc…"`：视图不替宿主跳转，跳书签是这里的决定
  v.on('click:element', ({ href }) => {
    if (href === undefined || editing.value) return;
    if (!href.startsWith('#')) {
      window.open(href, '_blank', 'noopener');
      return;
    }
    const target = d.rangeOfBookmark(href.slice(1));
    if (target !== undefined) v.scrollTo(target, { align: 'start', behavior: 'smooth' });
  });
  runSearch();
}

// 查找 = doc.find（模型）→ view.decorate（视图）→ view.scrollTo，api.md §15 的第一条配方
let ranges: DocRange[] = [];
let marks: DecorationHandle[] = [];

function clearSearch(): void {
  for (const m of marks) m.dispose();
  marks = [];
  ranges = [];
  hits.value = { total: 0, at: -1 };
}

function runSearch(): void {
  clearSearch();
  const d = doc.value;
  const v = view.value;
  if (d === undefined || v === undefined || query.value === '') return;
  ranges = d.find(query.value, { limit: 500 });
  marks = ranges.map((r) => v.decorate(r, { className: 'uwd-hit' }));
  hits.value = { total: ranges.length, at: -1 };
}

function step(delta: number): void {
  const v = view.value;
  if (v === undefined || ranges.length === 0) return;
  const { at } = hits.value;
  if (at >= 0) {
    marks[at]?.dispose();
    marks[at] = v.decorate(ranges[at] as DocRange, { className: 'uwd-hit' });
  }
  const next = (at + delta + ranges.length) % ranges.length;
  marks[next]?.dispose();
  marks[next] = v.decorate(ranges[next] as DocRange, { className: 'uwd-hit uwd-hit-current' });
  v.scrollTo(ranges[next] as DocRange, { align: 'center', behavior: 'smooth' });
  hits.value = { total: ranges.length, at: next };
}

function fitWidth(): void {
  const v = view.value;
  if (v === undefined) return;
  v.setZoom('fit-width');
  zoom.value = Math.round(v.zoom * 100);
}

async function exportDocx(): Promise<void> {
  const d = doc.value;
  if (d === undefined) return;
  const url = URL.createObjectURL(await d.toDocx());
  const a = document.createElement('a');
  a.href = url;
  // 加 -uw 后缀，免得覆盖原件 —— 与原件并排打开对比是导出的主要用法
  a.download = `${fileName.value.replace(/\.docx$/i, '')}-uw.docx`;
  a.click();
  URL.revokeObjectURL(url);
}

function pickFile(event: Event): void {
  const file = (event.target as HTMLInputElement).files?.[0];
  if (file !== undefined) void open(file, file.name, 'file');
  (event.target as HTMLInputElement).value = '';
}

function onDrop(event: DragEvent): void {
  dragging.value = false;
  const file = event.dataTransfer?.files[0];
  if (file !== undefined) void open(file, file.name, 'file');
}

function openSample(sample: Sample): void {
  void open(sample.url, `${sample.id}.docx`, sample.id);
}

function teardown(): void {
  clearSearch();
  for (const e of docEvents) e.dispose();
  docEvents = [];
  view.value?.dispose();
  view.value = undefined;
}

watch(zoom, (z) => {
  if (view.value !== undefined && Math.round(view.value.zoom * 100) !== z) view.value.setZoom(z / 100);
});
watch(editing, () => mount(false));

onMounted(async () => {
  const mod = await import('ultimate-word');
  api.value = mod.UltimateWord;
  const first = samples[0];
  if (first !== undefined) openSample(first);
});

onBeforeUnmount(teardown);
</script>

<template>
  <section
    class="uwd"
    :class="{ 'is-dragging': dragging }"
    @dragover.prevent="dragging = true"
    @dragleave.self="dragging = false"
    @drop.prevent="onDrop"
  >
    <div class="uwd-sources">
      <div class="uwd-samples" role="radiogroup" aria-label="样本文档">
        <button
          v-for="s in samples"
          :key="s.id"
          type="button"
          role="radio"
          class="uwd-sample"
          :aria-checked="current === s.id"
          :title="s.hint"
          :disabled="loading"
          @click="openSample(s)"
        >
          {{ s.label }}
        </button>
      </div>
      <label class="uwd-open">
        打开自己的 docx
        <input type="file" accept=".docx" hidden @change="pickFile" />
      </label>
      <span v-if="current === 'file'" class="uwd-filename" :title="fileName">{{ fileName }}</span>
    </div>

    <div class="uwd-toolbar">
      <label class="uwd-zoom">
        <span>缩放</span>
        <input v-model.number="zoom" type="range" min="30" max="300" step="10" :disabled="!doc" />
        <output>{{ zoom }}%</output>
      </label>
      <button type="button" class="uwd-btn" :disabled="!doc" @click="fitWidth">适应宽度</button>
      <label class="uwd-switch">
        <input v-model="editing" type="checkbox" :disabled="!doc" />
        <span>编辑</span>
      </label>
      <label class="uwd-find">
        <span>查找</span>
        <input
          v-model="query"
          type="search"
          placeholder="回车跳到下一处"
          :disabled="!doc"
          @input="runSearch"
          @keydown.enter.prevent="step($event.shiftKey ? -1 : 1)"
        />
        <output v-if="query">
          {{ hits.total === 0 ? '没有命中' : hits.at < 0 ? `${hits.total} 处` : `${hits.at + 1} / ${hits.total}` }}
        </output>
      </label>
      <button type="button" class="uwd-btn uwd-export" :disabled="!doc" @click="exportDocx">导出 docx</button>
    </div>

    <div class="uwd-stage-wrap">
      <div ref="stage" class="uwd-stage" />
      <div v-if="loading && !doc" class="uwd-veil">正在解析和排版…</div>
      <div v-if="dragging" class="uwd-veil uwd-drop">松开即可打开这份 docx</div>
    </div>

    <dl v-if="stats" class="uwd-stats">
      <div><dt>页数</dt><dd>{{ stats.pages }}</dd></div>
      <div><dt>纸张</dt><dd>{{ stats.size }}</dd></div>
      <div><dt>解析与排版</dt><dd>{{ stats.ms.toFixed(1) }} ms</dd></div>
      <div :title="stats.diagnostics.slice(0, 8).join('\n')">
        <dt>诊断</dt><dd>{{ stats.diagnostics.length === 0 ? '无' : `${stats.diagnostics.length} 条` }}</dd>
      </div>
    </dl>
    <p v-if="error" class="uwd-error" role="alert">{{ error }}</p>
  </section>
</template>

<style>
/*
 * 不加 scoped：页面、装饰这些节点是库挂进来的，没有 Vue 的作用域属性。
 * 选择器一律以 .uwd 开头，免得漏到文档站别处。
 */
.uwd {
  --uwd-accent: #b8191f;
  --uwd-desk: #5f646b;
  --uwd-gap: 12px;
  display: grid;
  gap: var(--uwd-gap);
  font-size: 14px;
}
.dark .uwd {
  --uwd-accent: #e5484d;
  --uwd-desk: #2c2f35;
}

.uwd-sources,
.uwd-toolbar {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 8px 16px;
}

.uwd-samples {
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
}
/* 当前样本用红头公文那道红线标出来：全站唯一一处强调色 */
.uwd-sample {
  padding: 6px 12px 5px;
  border-bottom: 2px solid transparent;
  color: var(--vp-c-text-2);
  font-weight: 500;
  transition: color 0.15s;
}
.uwd-sample:hover:not(:disabled) {
  color: var(--vp-c-text-1);
}
.uwd-sample[aria-checked='true'] {
  color: var(--vp-c-text-1);
  border-bottom-color: var(--uwd-accent);
}
.uwd-sample:disabled {
  cursor: progress;
}

.uwd-open,
.uwd-btn {
  display: inline-flex;
  align-items: center;
  padding: 4px 12px;
  border: 1px solid var(--vp-c-divider);
  border-radius: 6px;
  background: var(--vp-c-bg-soft);
  color: var(--vp-c-text-1);
  cursor: pointer;
  transition: border-color 0.15s;
}
.uwd-open:hover,
.uwd-btn:hover:not(:disabled) {
  border-color: var(--vp-c-brand-1);
}
.uwd-btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
.uwd-open:focus-within,
.uwd-btn:focus-visible,
.uwd-sample:focus-visible {
  outline: 2px solid var(--vp-c-brand-1);
  outline-offset: 2px;
}
.uwd-filename {
  max-width: 24ch;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  color: var(--vp-c-text-2);
}

.uwd-toolbar {
  padding: 8px 12px;
  border: 1px solid var(--vp-c-divider);
  border-radius: 8px;
  background: var(--vp-c-bg-soft);
}
.uwd-toolbar label {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  color: var(--vp-c-text-2);
}
.uwd-toolbar output {
  color: var(--vp-c-text-1);
  font-variant-numeric: tabular-nums;
}
.uwd-zoom output {
  min-width: 4ch;
}
.uwd-zoom input {
  width: 120px;
  accent-color: var(--vp-c-brand-1);
}
.uwd-switch input {
  accent-color: var(--vp-c-brand-1);
}
.uwd-find input {
  width: 150px;
  padding: 3px 8px;
  border: 1px solid var(--vp-c-divider);
  border-radius: 6px;
  background: var(--vp-c-bg);
  color: var(--vp-c-text-1);
}
.uwd-find input:focus {
  border-color: var(--vp-c-brand-1);
  outline: none;
}
.uwd-export {
  margin-left: auto;
}

.uwd-stage-wrap {
  position: relative;
  border-radius: 8px;
  overflow: hidden;
  background: var(--uwd-desk);
}
.uwd-stage {
  height: min(78vh, 960px);
  overflow: auto;
  padding: 24px;
}
.uwd-veil {
  position: absolute;
  inset: 0;
  display: grid;
  place-items: center;
  color: #fff;
  font-size: 15px;
  background: rgb(0 0 0 / 0.35);
}
.uwd-drop {
  outline: 2px dashed rgb(255 255 255 / 0.8);
  outline-offset: -12px;
  background: rgb(0 0 0 / 0.5);
}

/* 页间距、纸张阴影是宿主的事，库只给位置（与调试台同一份约定） */
.uwd .uw-doc {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 24px;
}
.uwd .uw-page-shell {
  background: #fff;
  box-shadow: 0 1px 3px rgb(0 0 0 / 0.3), 0 8px 24px rgb(0 0 0 / 0.25);
}
.uwd .uw-page-filler {
  opacity: 0.6;
}
.uwd-hit {
  background: rgb(255 211 61 / 0.45);
}
.uwd-hit-current {
  background: rgb(255 170 60 / 0.55);
  outline: 1.5px solid var(--uwd-accent);
}

.uwd-stats {
  display: flex;
  flex-wrap: wrap;
  gap: 4px 28px;
  margin: 0;
  color: var(--vp-c-text-2);
  font-size: 13px;
}
.uwd-stats > div {
  display: flex;
  gap: 8px;
}
.uwd-stats dt {
  color: var(--vp-c-text-3);
}
.uwd-stats dd {
  margin: 0;
  color: var(--vp-c-text-1);
  font-variant-numeric: tabular-nums;
}
.uwd-error {
  margin: 0;
  color: var(--vp-c-danger-1);
}

@media (max-width: 640px) {
  .uwd-stage {
    height: 70vh;
    padding: 12px;
  }
  .uwd-export {
    margin-left: 0;
  }
}
@media (prefers-reduced-motion: reduce) {
  .uwd-sample,
  .uwd-open,
  .uwd-btn {
    transition: none;
  }
}
</style>
