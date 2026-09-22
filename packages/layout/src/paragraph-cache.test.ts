import { readFileSync } from 'node:fs';
import { createDiagnosticSink } from '@uw/core';
import { createTextMeasurer, FontRegistry } from '@uw/fonts';
import { loadBundledPacks } from '@uw/fonts/node';
import {
  createTextEditor,
  DEFAULT_SETTINGS,
  fontNameCandidates,
  loadDocument,
  resolveBody,
  walkParagraphs,
} from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import { describe, expect, it, vi } from 'vitest';
import { layoutDocumentWithFields } from './fields.ts';
import { WIDTH_RULES } from './items.ts';
import { OBJECT_RULES, SCRIPT_RULES } from './line-height.ts';
import type { LayoutParagraphOptions } from './paragraph.ts';
import { layoutParagraph } from './paragraph.ts';
import { ParagraphLayoutCache } from './paragraph-cache.ts';
import { fakeMeasurer, NO_GRID, para, run } from './test-fixtures.ts';

function options(): LayoutParagraphOptions {
  return {
    measurer: { ...fakeMeasurer(), revision: 0 },
    contentWidth: 2100,
    settings: structuredClone(DEFAULT_SETTINGS),
    docGrid: NO_GRID,
    paragraphCache: new ParagraphLayoutCache(),
  };
}

function required<T>(value: T | undefined): T {
  if (value === undefined) throw new Error('测试缺少预期节点');
  return value;
}

describe('段落缓存', () => {
  it('重建的同值段落直接复用，调用方修改布局不会污染下一帧', () => {
    const opts = options();
    const measured = vi.spyOn(opts.measurer, 'advances');
    const p = para([run('正文中文与 ABC')]);
    const first = layoutParagraph(p, opts);
    const expected = structuredClone(first);
    expect(measured).toHaveBeenCalled();
    measured.mockClear();
    required(required(first.lines[0]).fragments[0]).text = '被调用方修改';
    const second = layoutParagraph(structuredClone(p), opts);
    expect(second).toEqual(expected);
    required(second.lines[0]).height = -1;
    expect(layoutParagraph(p, opts)).toEqual(expected);
    expect(measured).not.toHaveBeenCalled();
    expect(structuredClone(second)).toEqual(second);
  });

  it.each(['width', 'grid', 'settings', 'font', 'objectRules', 'scriptRules', 'widthRules'])(
    '%s 变化必须重算并与无缓存结果一致',
    (kind) => {
      const opts = options();
      const p = para([run('中文 A 中文'.repeat(8))]);
      layoutParagraph(p, opts);
      const measured = vi.spyOn(opts.measurer, 'advances');
      if (kind === 'width') opts.contentWidth = 1200;
      if (kind === 'grid') opts.docGrid = { ...NO_GRID, type: 'lines', linePitch: 400 };
      if (kind === 'settings') opts.settings.defaultTabStop += 100;
      if (kind === 'font') opts.defaultFont = 'another';
      if (kind === 'objectRules') opts.objectRules = { ...OBJECT_RULES };
      if (kind === 'scriptRules') opts.scriptRules = { ...SCRIPT_RULES };
      if (kind === 'widthRules') opts.widthRules = { ...WIDTH_RULES };
      const actual = layoutParagraph(p, opts);
      expect(measured).toHaveBeenCalled();
      const { paragraphCache: _, ...uncached } = opts;
      expect(actual).toEqual(layoutParagraph(p, uncached));
    },
  );

  it('原地修改文字和样式也失效，不能只看节点引用或 id', () => {
    const opts = options();
    const p = para([run('中文')]);
    const { paragraphCache: _, ...uncached } = opts;
    layoutParagraph(p, opts);
    for (const update of [
      () => {
        required(p.runs[0]).content = [{ kind: 'text', text: '修改后的文字'.repeat(5) }];
      },
      () => {
        required(p.runs[0]).props.size = 320;
      },
      () => {
        p.props.justification = 'center';
      },
    ]) {
      update();
      expect(layoutParagraph(p, opts)).toEqual(layoutParagraph(p, uncached));
    }
  });

  it('只由本段域显示值失效，空字符串与未求值不同', () => {
    const opts = options();
    const r = run('旧域值');
    const p = para([r]);
    const values = new Map<string, string>();
    opts.fieldValues = values;
    layoutParagraph(p, opts);
    const measured = vi.spyOn(opts.measurer, 'advances');
    values.set('其他段落', '10');
    layoutParagraph(p, opts);
    expect(measured).not.toHaveBeenCalled();
    const { paragraphCache: _, ...uncached } = opts;
    for (const value of ['', '1', '100']) {
      values.set(r.id, value);
      expect(layoutParagraph(p, opts)).toEqual(layoutParagraph(p, uncached));
    }
  });

  it('度量器或版本变化时清空；无版本的自定义度量器每次重算', () => {
    const opts = options();
    const p = para([run('中文')]);
    let revision = 0;
    opts.measurer = {
      ...fakeMeasurer(),
      get revision() {
        return revision;
      },
    };
    layoutParagraph(p, opts);
    const measured = vi.spyOn(opts.measurer, 'advances');
    revision++;
    layoutParagraph(p, opts);
    expect(measured).toHaveBeenCalled();
    expect(opts.paragraphCache?.size).toBe(1);
    opts.measurer = { ...fakeMeasurer(), revision };
    const replacement = vi.spyOn(opts.measurer, 'advances');
    layoutParagraph(p, opts);
    expect(replacement).toHaveBeenCalled();
    opts.measurer = fakeMeasurer();
    const unversioned = vi.spyOn(opts.measurer, 'advances');
    layoutParagraph(p, opts);
    unversioned.mockClear();
    layoutParagraph(p, opts);
    expect(unversioned).toHaveBeenCalled();
    expect(opts.paragraphCache?.size).toBe(0);
  });

  it('LRU 容量有界，命中会保留热项，clear 可显式释放', () => {
    const opts = options();
    const cache = new ParagraphLayoutCache(2);
    opts.paragraphCache = cache;
    const [a, b, c] = [para([run('甲')]), para([run('乙')]), para([run('丙')])] as const;
    for (const p of [a, b, a, c]) layoutParagraph(p, opts);
    expect(cache.size).toBe(2);
    const measured = vi.spyOn(opts.measurer, 'advances');
    layoutParagraph(a, opts);
    expect(measured).not.toHaveBeenCalled();
    layoutParagraph(b, opts);
    expect(measured).toHaveBeenCalled();
    cache.clear();
    expect(cache.size).toBe(0);
    expect(() => new ParagraphLayoutCache(0)).toThrow(RangeError);
  });
});

function fixture(name: string) {
  const sink = createDiagnosticSink();
  const bytes = readFileSync(new URL(`../../../apps/fidelity/fixtures/${name}.docx`, import.meta.url));
  const loaded = loadDocument(OpcPackage.open(bytes), sink);
  const registry = new FontRegistry();
  for (const pack of loadBundledPacks()) registry.registerMetrics(pack);
  const measurer = createTextMeasurer(registry, {
    candidates: (family) => fontNameCandidates(loaded.fonts, family),
  });
  const opts = { measurer, settings: loaded.cascade.settings, headerFooters: loaded.headerFooters };
  return { loaded, opts };
}

describe('缓存与完整排版一致', () => {
  it.each([
    'gongwen-01',
    'spike-header-03',
    'spike-image-03',
    'spike-table-04',
    'spike-page-02',
    'spike-width-01',
  ])('%s 的页、行、坐标和域求值保持一致，第二次不再度量', (name) => {
    const { loaded, opts } = fixture(name);
    const expected = layoutDocumentWithFields(loaded.resolved, loaded.fields, opts);
    const cached = { ...opts, paragraphCache: new ParagraphLayoutCache() };
    expect(layoutDocumentWithFields(loaded.resolved, loaded.fields, cached)).toEqual(expected);
    const measured = vi.spyOn(opts.measurer, 'advances');
    expect(layoutDocumentWithFields(structuredClone(loaded.resolved), loaded.fields, cached)).toEqual(
      expected,
    );
    expect(measured).not.toHaveBeenCalled();
  });

  it('编辑、拆段、撤销和重做都与完整重排一致，只度量变化段落', () => {
    const { loaded, opts } = fixture('gongwen-01');
    const editor = createTextEditor(loaded.body);
    const cached = { ...opts, paragraphCache: new ParagraphLayoutCache() };
    const initial = layoutDocumentWithFields(loaded.resolved, loaded.fields, cached);
    const p = required(
      [...walkParagraphs(editor.body)].find((p) => p.runs.some((r) => r.content[0]?.kind === 'text')),
    );
    const r = required(p.runs.find((r) => r.content[0]?.kind === 'text'));
    const pos = { nodeId: r.id, contentIndex: 0, offset: 0 };
    const measured = vi.spyOn(opts.measurer, 'advances');
    const check = () => {
      const resolved = resolveBody(loaded.cascade, editor.body);
      measured.mockClear();
      const actual = layoutDocumentWithFields(resolved, loaded.fields, cached);
      const calls = measured.mock.calls.length;
      measured.mockClear();
      expect(actual).toEqual(layoutDocumentWithFields(resolved, loaded.fields, opts));
      expect(calls).toBeLessThan(measured.mock.calls.length);
      return actual;
    };
    editor.tx((tx) => {
      tx.insertText(pos, '新增中文段落内容'.repeat(80));
    });
    check();
    editor.tx((tx) => {
      tx.splitParagraph({ ...pos, offset: 10 });
    });
    check();
    editor.undo();
    check();
    editor.undo();
    expect(check()).toEqual(initial);
    editor.redo();
    check();
  });
});
