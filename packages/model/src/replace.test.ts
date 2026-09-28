/**
 * 查找替换：`tx.replaceText` 的格式继承、`findMatches` 的捕获组、`expandReplacement` 的模板规则、
 * `rangeEditable` 对域的判断。格式是这里最容易错的一处 —— 先删后插会继承左边的 run，
 * 命中恰好从 run 开头起时新文字就丢了粗体，屏幕上看不出是「替换」造成的。
 */
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { CascadeContext } from './cascade.ts';
import type { Body, Paragraph } from './nodes.ts';
import { paragraphText, walkParagraphs } from './nodes.ts';
import { EMPTY_NUMBERING } from './numbering.ts';
import { parseBody } from './parse-body.ts';
import type { DocRange } from './position.ts';
import { resolveBody } from './resolve-body.ts';
import type { TextMatch } from './search.ts';
import { expandReplacement, findMatches, findText } from './search.ts';
import { DEFAULT_SETTINGS } from './settings.ts';
import { parseStyles } from './styles.ts';
import { createTextEditor, rangeEditable } from './text-transaction.ts';
import { EMPTY_THEME } from './theme.ts';

const CTX: CascadeContext = {
  styles: parseStyles(parseXml('<w:styles/>'), createDiagnosticSink()),
  theme: EMPTY_THEME,
  settings: DEFAULT_SETTINGS,
  numbering: EMPTY_NUMBERING,
};

function parse(xml: string): Body {
  return parseBody(parseXml(`<w:document><w:body>${xml}</w:body></w:document>`), createDiagnosticSink());
}
function para(body: Body, index = 0): Paragraph {
  return [...walkParagraphs(body)][index] as Paragraph;
}
function find(body: Body, pattern: string | RegExp): DocRange[] {
  return findText(resolveBody(CTX, body), pattern);
}
/** 每个 run 的文字与是否加粗，删空的 run 不列 —— 删除保留槽位，空 run 不是这里要比的 */
function runs(body: Body, index = 0): [string, boolean][] {
  return para(body, index)
    .runs.map((r): [string, boolean] => [
      r.content.map((c) => (c.kind === 'text' ? c.text : c.kind === 'tab' ? '\t' : '\n')).join(''),
      r.props.bold === true,
    ])
    .filter(([t]) => t !== '');
}

const BOLD = '<w:rPr><w:b/></w:rPr>';

describe('tx.replaceText', () => {
  it('新文字取命中首字的格式，即使首字在 run 开头', () => {
    const body = parse(
      `<w:p><w:r><w:t>请</w:t></w:r><w:r>${BOLD}<w:t>签</w:t></w:r><w:r><w:t>发人签字</w:t></w:r></w:p>`,
    );
    const editor = createTextEditor(body);
    const before = editor.body;
    const [hit] = find(body, '签发人');
    let out: DocRange | undefined;
    editor.tx((t) => {
      out = t.replaceText(hit as DocRange, '审批人');
    });
    expect(paragraphText(para(editor.body))).toBe('请审批人签字');
    expect(runs(editor.body)).toEqual([
      ['请', false],
      ['审批人', true],
      ['签字', false],
    ]);
    // 返回的范围正好框住新文字，可以直接拿去选中
    expect(out).toEqual({
      start: { ...(hit as DocRange).start },
      end: { ...(hit as DocRange).start, offset: 3 },
    });
    editor.undo();
    expect(editor.body).toBe(before);
  });

  it('同一 run 中间的命中、代理对首字、空替换', () => {
    const body = parse('<w:p><w:r><w:t>甲乙丙丁😀戊</w:t></w:r></w:p>');
    const editor = createTextEditor(body);
    const [a] = find(body, '乙丙');
    const [b] = find(body, '😀戊');
    editor.tx((t) => {
      // 后面的先换：前面的位置不受影响
      t.replaceText(b as DocRange, '');
      const r = t.replaceText(a as DocRange, 'XYZ');
      expect(r.end.offset - r.start.offset).toBe(3);
    });
    expect(paragraphText(para(editor.body))).toBe('甲XYZ丁');
  });

  it('制表位与换行写成真的片段，返回的终点在最后一个片段后', () => {
    const body = parse(`<w:p><w:r>${BOLD}<w:t>姓名：张三</w:t></w:r></w:p>`);
    const editor = createTextEditor(body);
    const [hit] = find(body, '：');
    editor.tx((t) => {
      const r = t.replaceText(hit as DocRange, '\t李\n四');
      expect(r.end.contentIndex).toBeGreaterThan(r.start.contentIndex);
    });
    expect(runs(editor.body)).toEqual([['姓名\t李\n四张三', true]]);
  });

  it('首字是制表位时退回先删后插，范围折叠时就是插入', () => {
    const body = parse('<w:p><w:r><w:t>甲</w:t><w:tab/><w:t>乙</w:t></w:r></w:p>');
    const editor = createTextEditor(body);
    const [hit] = find(body, '\t乙');
    editor.tx((t) => {
      const r = t.replaceText(hit as DocRange, '丙');
      t.replaceText({ start: r.end, end: r.end }, '丁');
    });
    expect(paragraphText(para(editor.body))).toBe('甲丙丁');
  });

  it('命中域的显示文字时事务整个回滚', () => {
    const body = parse(
      '<w:p><w:r><w:t>第</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>页</w:t></w:r></w:p>',
    );
    const editor = createTextEditor(body);
    const before = editor.body;
    const [inField] = find(body, '1');
    const [across] = find(body, '第1页');
    const [plain] = find(body, '页');
    expect(rangeEditable(body, inField as DocRange)).toBe(false);
    expect(rangeEditable(body, across as DocRange)).toBe(false);
    expect(rangeEditable(body, plain as DocRange)).toBe(true);
    expect(() =>
      editor.tx((t) => {
        t.replaceText(plain as DocRange, '頁');
        t.replaceText(inField as DocRange, '2');
      }),
    ).toThrow();
    expect(editor.body).toBe(before);
  });
});

describe('findMatches / expandReplacement', () => {
  const body = resolveBody(CTX, parse('<w:p><w:r><w:t>2026-09-28 与 2025-01-02</w:t></w:r></w:p>'));
  const re = /(?<y>\d{4})-(\d{2})-(\d{2})/;

  it('带着捕获组、命中文字与本段的串', () => {
    const [m] = findMatches(body, re);
    expect(m?.text).toBe('2026-09-28');
    expect(m?.captures).toEqual(['2026-09-28', '2026', '09', '28']);
    expect(m?.groups).toEqual({ y: '2026' });
    expect(m?.input).toBe('2026-09-28 与 2025-01-02');
    expect(findMatches(body, re).map((x) => x.range)).toEqual(findText(body, re));
  });

  it('模板规则照 String.prototype.replace', () => {
    // 非全局正则的 replace 只换第一处，拿第一个命中比
    const m = findMatches(body, re)[0] as TextMatch;
    for (const tpl of ['$<y>年$2月$3日', '[$&]', '$$1', '$4$0$', "$`|$'", '$10', '$<nope>', '$<y']) {
      const got =
        m.input.slice(0, m.index) + expandReplacement(tpl, m) + m.input.slice(m.index + m.text.length);
      expect(got, tpl).toBe(m.input.replace(re, tpl));
    }
  });

  it('没有命名组时 $< 原样留着，前后断言看得见上下文', () => {
    const [m] = findMatches(body, /(?<=与 )\d+/);
    expect(m?.text).toBe('2025');
    expect(expandReplacement('$<x>[$&]', m as TextMatch)).toBe('$<x>[2025]');
  });
});
