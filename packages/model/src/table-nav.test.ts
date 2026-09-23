import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { Body } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { parseBody } from './parse-body.ts';
import { adjacentCellRange } from './table-nav.ts';

function parse(xml: string): Body {
  return parseBody(parseXml(`<w:document><w:body>${xml}</w:body></w:document>`), createDiagnosticSink());
}
const p = (text: string) => (text ? `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>` : '<w:p/>');
const tc = (inner: string, pr = '') => `<w:tc>${pr ? `<w:tcPr>${pr}</w:tcPr>` : ''}${inner}</w:tc>`;
const tr = (...cells: string[]) => `<w:tr>${cells.join('')}</w:tr>`;
const tbl = (...rows: string[]) => `<w:tbl>${rows.join('')}</w:tbl>`;

describe('表格里的 Tab 导航', () => {
  const body = parse(
    p('正文') +
      tbl(
        tr(tc(p('甲') + p('甲二'), '<w:vMerge w:val="restart"/>'), tc(p('乙'))),
        tr(tc(p(''), '<w:vMerge/>'), tc(p('') + tbl(tr(tc(p('内一')), tc(p('内二')))))),
      ),
  );
  const paragraphs = [...walkParagraphs(body)];
  const id = (text: string) => {
    const found = paragraphs.find((x) =>
      x.runs.some((r) => r.content.some((c) => c.kind === 'text' && c.text === text)),
    );
    return found?.runs[0]?.id ?? '';
  };
  const empty = paragraphs.filter((x) => !x.runs.length).map((x) => x.id);

  it('不在单元格里答 undefined，行末接下一行，选中整格（多段）', () => {
    expect(adjacentCellRange(body, id('正文'), 'forward')).toBeUndefined();
    expect(adjacentCellRange(body, id('甲'), 'forward')?.start.nodeId).toBe(id('乙'));
    const back = adjacentCellRange(body, id('乙'), 'backward');
    expect([back?.start.nodeId, back?.end.nodeId]).toEqual([id('甲'), id('甲二')]);
  });

  it('跳过 vMerge=continue；含嵌套表的格从它自己的首段选到内表末段', () => {
    const next = adjacentCellRange(body, id('乙'), 'forward');
    expect([next?.start.nodeId, next?.end.nodeId]).toEqual([empty[1], id('内二')]);
    // 光标落在续格里（点不到，但位置合法）也照样往后走
    expect(adjacentCellRange(body, empty[0] ?? '', 'forward')?.start.nodeId).toBe(empty[1]);
  });

  it('嵌套表在内表里走，内表首末格答 null，不跳回外表', () => {
    expect(adjacentCellRange(body, id('内一'), 'forward')?.start.nodeId).toBe(id('内二'));
    expect(adjacentCellRange(body, id('内二'), 'forward')).toBeNull();
    expect(adjacentCellRange(body, id('内一'), 'backward')).toBeNull();
    expect(adjacentCellRange(body, id('甲'), 'backward')).toBeNull();
  });
});
