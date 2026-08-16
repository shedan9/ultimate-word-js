/**
 * 域的结构还原。
 *
 * 两半分开测：`parseFieldInstruction` 是纯字符串函数，`scanFields` 要一棵真的树。
 * 指令那一半的期望值来自 ECMA-376 §17.16 与 Word 的实际写法（`\* MERGEFORMAT` 那些），
 * 不是拍脑袋 —— 但「开关后面那个词是它的值」是启发式，见 fields.ts 里写明的简化。
 */
import type { DiagnosticSink } from '@uw/core';
import { createDiagnosticSink } from '@uw/core';
import { parseXml } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import type { CascadeContext } from './cascade.ts';
import { fieldHyperlinks, fieldSwitch, parseFieldInstruction, scanFields } from './fields.ts';
import type { Body } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import { EMPTY_NUMBERING } from './numbering.ts';
import { parseBody } from './parse-body.ts';
import { resolveBody } from './resolve-body.ts';
import { DEFAULT_SETTINGS } from './settings.ts';
import { parseStyles } from './styles.ts';
import { EMPTY_THEME } from './theme.ts';

/** 级联在这个文件里不是主角，一份空样式表就够 */
const EMPTY_CTX: CascadeContext = {
  styles: parseStyles(parseXml('<w:styles/>'), createDiagnosticSink()),
  theme: EMPTY_THEME,
  settings: DEFAULT_SETTINGS,
  numbering: EMPTY_NUMBERING,
};

const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';

function parse(bodyXml: string): { body: Body; sink: DiagnosticSink } {
  const sink = createDiagnosticSink();
  const doc = parseXml(`<w:document><w:body>${bodyXml}${SECT}</w:body></w:document>`, 'document.xml');
  return { body: parseBody(doc, sink), sink };
}

/** 复杂域的三段：begin + 指令 + separate + 结果 + end */
function field(instr: string, result: string): string {
  return `<w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>${result}</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>`;
}

describe('指令分词', () => {
  it('第一个词是域类型，大小写归一', () => {
    expect(parseFieldInstruction(' page ').type).toBe('PAGE');
    expect(parseFieldInstruction('Ref _Ref1').type).toBe('REF');
  });

  it('引号剥掉、引号里的空格保住', () => {
    const f = parseFieldInstruction('HYPERLINK "http://a.b/c d" ');
    expect(f.args).toEqual(['http://a.b/c d']);
  });

  it('引号里的 \\" 是引号本身', () => {
    expect(parseFieldInstruction('TITLE "他说\\"好\\""').args).toEqual(['他说"好"']);
  });

  it('引号外的反斜杠原样留着 —— 那是路径不是转义', () => {
    const f = parseFieldInstruction('INCLUDEPICTURE C:\\pic\\a.png');
    expect(f.args).toEqual(['C:\\pic\\a.png']);
    expect(f.switches).toEqual([]);
  });

  it('开关取 \\ 后面那一个字符，后面那个词是它的值', () => {
    const f = parseFieldInstruction('TOC \\o "1-3" \\h \\z \\u');
    expect(f.switches).toEqual([{ name: 'o', value: '1-3' }, { name: 'h' }, { name: 'z' }, { name: 'u' }]);
  });

  it('\\* MERGEFORMAT：星号也是开关名', () => {
    const f = parseFieldInstruction('PAGE \\* MERGEFORMAT');
    expect(f.type).toBe('PAGE');
    expect(fieldSwitch(f, '*')?.value).toBe('MERGEFORMAT');
  });

  it('不带空格的 \\o"1-3" 也认', () => {
    expect(fieldSwitch(parseFieldInstruction('TOC \\o"1-3"'), 'o')?.value).toBe('1-3');
  });

  it('开关查找不分大小写', () => {
    expect(fieldSwitch(parseFieldInstruction('REF a \\H'), 'h')).toEqual({ name: 'H' });
  });

  it('空指令不炸，给一个空类型', () => {
    expect(parseFieldInstruction('   ')).toEqual({ type: '', args: [], switches: [] });
  });
});

describe('扫描：复杂域', () => {
  it('界桩配成一个域，指令与结果各就各位', () => {
    const { body } = parse(`<w:p>${field(' PAGE \\* MERGEFORMAT ', '3')}</w:p>`);
    const [f] = scanFields(body);
    expect(f?.kind).toBe('complex');
    expect(f?.instructionText).toBe(' PAGE \\* MERGEFORMAT ');
    expect(f?.instruction.type).toBe('PAGE');
    expect(f?.depth).toBe(0);
    expect(f?.resultRuns).toHaveLength(1);
    expect(f?.begin?.contentIndex).toBe(0);
  });

  it('切碎的 instrText 接回一条完整指令', () => {
    // Word 存盘时常把一条指令拆成好几段，拼接**不加分隔符**，加了会把域名割开
    const { body } = parse(
      `<w:p>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> HYPER</w:instrText></w:r>
        <w:r><w:instrText xml:space="preserve">LINK "http://x" </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>x</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>`,
    );
    expect(scanFields(body)[0]?.instruction.type).toBe('HYPERLINK');
  });

  it('域跨段落也能配上 —— TOC 就是这么长的', () => {
    const { body } = parse(
      `<w:p>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> TOC \\o "1-3" </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      </w:p>
      <w:p><w:r><w:t>第一章</w:t></w:r></w:p>
      <w:p><w:r><w:t>第二章</w:t></w:r></w:p>
      <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
    );
    const [f] = scanFields(body);
    expect(f?.instruction.type).toBe('TOC');
    expect(f?.resultRuns).toHaveLength(2);
    // begin 与 end 落在不同段落上 —— 这正是「按段落扫永远配不上」的那种情况
    expect(f?.begin?.paragraphId).not.toBe(f?.end?.paragraphId);
  });

  it('没有 separate 的域不显示任何东西', () => {
    const { body } = parse(
      `<w:p>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> PAGE </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>`,
    );
    const [f] = scanFields(body);
    expect(f?.separate).toBeUndefined();
    expect(f?.resultRuns).toEqual([]);
  });

  it('嵌套：内层指令归内层，外层因此缺一块（求值期才补）', () => {
    const { body } = parse(
      `<w:p>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> IF </w:instrText></w:r>
        ${field(' PAGE ', '1')}
        <w:r><w:instrText> = 1 "首页" "" </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>首页</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>`,
    );
    const fields = scanFields(body);
    // 顺序按 begin 排：外层在前，尽管它是后结束的
    expect(fields.map((f) => f.instruction.type)).toEqual(['IF', 'PAGE']);
    expect(fields[0]?.depth).toBe(0);
    expect(fields[1]?.depth).toBe(1);
    expect(fields[0]?.instructionText).toBe(' IF  = 1 "首页" "" ');
    expect(fields[1]?.instructionText).toBe(' PAGE ');
  });

  it('多出来的 end 记诊断后继续，不抛', () => {
    const { body } = parse(
      `<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>正文</w:t></w:r></w:p>`,
    );
    const sink = createDiagnosticSink();
    expect(scanFields(body, sink)).toEqual([]);
    expect(sink.list()[0]?.code).toBe('field-unbalanced');
  });

  it('少了 end 的域照样吐出来，结果区为空', () => {
    const { body } = parse(
      `<w:p>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> PAGE </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>3</w:t></w:r>
      </w:p>`,
    );
    const sink = createDiagnosticSink();
    const [f] = scanFields(body, sink);
    expect(f?.end).toBeUndefined();
    expect(f?.resultRuns).toEqual([]);
    expect(sink.list()[0]?.code).toBe('field-unclosed');
  });
});

describe('扫描：简单域（w:fldSimple）', () => {
  it('压缩写法也收成一个域', () => {
    const { body } = parse(`<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p>`);
    const [f] = scanFields(body);
    expect(f?.kind).toBe('simple');
    expect(f?.instruction.type).toBe('PAGE');
    expect(f?.resultRuns).toHaveLength(1);
    expect(f?.begin).toBeUndefined();
  });

  it('挨着的两个同指令简单域是两个域，不能并成一个', () => {
    const { body } = parse(
      `<w:p>
        <w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>
        <w:fldSimple w:instr="PAGE"><w:r><w:t>2</w:t></w:r></w:fldSimple>
      </w:p>`,
    );
    const fields = scanFields(body);
    expect(fields).toHaveLength(2);
    expect(fields[0]?.resultRuns).toHaveLength(1);
  });
});

describe('HYPERLINK 落到 run 上', () => {
  it('地址来自指令，结果 run 全都带上链接', () => {
    const { body } = parse(`<w:p>${field(' HYPERLINK "http://a.b/c" ', '点这里')}</w:p>`);
    const links = fieldHyperlinks(scanFields(body));
    const ids = scanFields(body)[0]?.resultRuns ?? [];
    expect(ids).toHaveLength(1);
    expect(links.get(ids[0] as string)).toEqual({ url: 'http://a.b/c' });
  });

  it('\\l 是书签名，没有 url 的纯文档内跳转也算数', () => {
    const { body } = parse(`<w:p>${field(' HYPERLINK \\l "_Toc1" ', '第一章')}</w:p>`);
    const f = scanFields(body)[0];
    const links = fieldHyperlinks([f as NonNullable<typeof f>]);
    expect(links.get((f?.resultRuns[0] ?? '') as string)).toEqual({ anchor: '_Toc1' });
  });

  it('界桩所在的 run 不算结果 —— 链接不会蔓延到域外', () => {
    const { body } = parse(
      `<w:p><w:r><w:t>前</w:t></w:r>${field(' HYPERLINK "http://a" ', '中')}<w:r><w:t>后</w:t></w:r></w:p>`,
    );
    const links = fieldHyperlinks(scanFields(body));
    expect(links.size).toBe(1);
  });

  it('不是 HYPERLINK 的域不产生链接', () => {
    const { body } = parse(`<w:p>${field(' PAGE ', '3')}</w:p>`);
    expect(fieldHyperlinks(scanFields(body)).size).toBe(0);
  });

  it('级联那一趟把链接铺到 resolved 树上 —— 渲染层不必认识「域」', () => {
    const { body } = parse(
      `<w:p>
        <w:hyperlink r:id="rId7"><w:r><w:t>容器</w:t></w:r></w:hyperlink>
        ${field(' HYPERLINK "http://a" ', '域')}
      </w:p>`,
    );
    const resolved = resolveBody(EMPTY_CTX, body, { hyperlinks: fieldHyperlinks(scanFields(body)) });
    const runs = [...walkParagraphs(resolved)][0]?.runs ?? [];
    // 六个 run：容器 / begin / 指令 / separate / 结果 / end
    expect(runs).toHaveLength(6);
    // 两条路最终落在同一个字段上：容器给 relId，域给 url
    expect(runs[0]?.hyperlink).toEqual({ relId: 'rId7' });
    expect(runs[4]?.hyperlink).toEqual({ url: 'http://a' });
    // 界桩那几个 run 不带链接，否则点在域代码上也会跳转
    expect(runs[1]?.hyperlink).toBeUndefined();
    expect(runs[5]?.hyperlink).toBeUndefined();
  });
});

describe('可结构化克隆（原则 1.1）', () => {
  it('扫描结果能过 Worker 边界', () => {
    const { body } = parse(`<w:p>${field(' PAGE ', '3')}</w:p>`);
    const fields = scanFields(body);
    expect(() => structuredClone(fields)).not.toThrow();
    expect(structuredClone(fields)).toEqual(fields);
  });
});
