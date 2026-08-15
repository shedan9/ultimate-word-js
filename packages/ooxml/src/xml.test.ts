/**
 * XML 编解码的回归测试。
 *
 * 重点全在「别把数据弄丢 / 弄变形」上 —— 这一层出一个静默的失真（把 `w:val="00"`
 * 读成数字 0、把 `xml:space="preserve"` 里的空格 trim 掉），故障会在几百行之外的
 * 排版结果里冒出来，查起来极贵。
 */
import { UwError } from '@uw/core';
import { describe, expect, it } from 'vitest';
import { attr, child, children, parseXml, serializeXml, textContent } from './xml.ts';

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

describe('parseXml', () => {
  it('保留元素顺序与重复元素', () => {
    const { root } = parseXml(`${DECL}<w:p><w:r>A</w:r><w:br/><w:r>B</w:r></w:p>`);
    expect(root.name).toBe('w:p');
    expect(root.children.map((c) => (c.kind === 'element' ? c.name : c.kind))).toEqual([
      'w:r',
      'w:br',
      'w:r',
    ]);
    expect(children(root, 'w:r').map(textContent)).toEqual(['A', 'B']);
  });

  it('属性一律保持字符串 —— w:val="00" 不能变成数字 0', () => {
    const { root } = parseXml('<w:b w:val="00" w:sz="0011"/>');
    expect(attr(root, 'w:val')).toBe('00');
    expect(attr(root, 'w:sz')).toBe('0011');
  });

  it('xml:space="preserve" 的首尾空格是正文，不能 trim', () => {
    const { root } = parseXml('<w:r><w:t xml:space="preserve">  张三 </w:t></w:r>');
    const t = child(root, 'w:t');
    expect(t && textContent(t)).toBe('  张三 ');
  });

  it('实体解码成字符', () => {
    const { root } = parseXml('<w:t>a &amp; b &lt;c&gt; &quot;d&quot;</w:t>');
    expect(textContent(root)).toBe('a & b <c> "d"');
  });

  it('数值字符引用也要解 —— 十进制与十六进制，文本与属性都算', () => {
    const { root } = parseXml('<w:t v="a&#9;b">&#x41;&#169;</w:t>');
    expect(textContent(root)).toBe('A©');
    expect(attr(root, 'v')).toBe('a\tb');
  });

  it('保留 XML 声明与注释', () => {
    const doc = parseXml(`${DECL}<w:p><!-- 别的工具留下的 --></w:p>`);
    expect(doc.declaration).toEqual({ version: '1.0', encoding: 'UTF-8', standalone: 'yes' });
    expect(doc.root.children[0]).toEqual({ kind: 'comment', text: ' 别的工具留下的 ' });
  });

  it('不认识的元素照样进树（原则 1.4 的地基）', () => {
    const { root } = parseXml('<w:p><mc:AlternateContent foo="1"><w:x/></mc:AlternateContent></w:p>');
    const alt = child(root, 'mc:AlternateContent');
    expect(alt?.attrs).toEqual({ foo: '1' });
    expect(alt && children(alt)).toHaveLength(1);
  });

  it('没有根元素时抛 UwError(MALFORMED_XML)', () => {
    expect(() => parseXml(DECL, '/word/document.xml')).toThrow(UwError);
    try {
      parseXml(DECL, '/word/document.xml');
    } catch (e) {
      expect((e as UwError).code).toBe('MALFORMED_XML');
      expect((e as UwError).part).toBe('/word/document.xml');
    }
  });
});

describe('serializeXml', () => {
  /** 判据是**语义等价**而不是逐字节相同 —— 空元素写法、可选转义都有多种合法形式 */
  const roundTrips = (xml: string): void => {
    const once = parseXml(xml);
    expect(parseXml(serializeXml(once))).toEqual(once);
  };

  it('往返后树不变', () => {
    roundTrips(`${DECL}<w:p w:rsidR="00A1"><w:r><w:t xml:space="preserve"> 甲 </w:t></w:r><w:br/></w:p>`);
    roundTrips('<a><!--c--><b x="1&quot;2"/>t &amp; t</a>');
  });

  it('转义文本里的 & 与 <', () => {
    const doc = parseXml('<w:t>a &amp; b</w:t>');
    expect(serializeXml(doc)).toBe('<w:t>a &amp; b</w:t>');
  });

  it('属性里的制表符转成数字实体 —— 否则被属性值规范化吃成空格', () => {
    const doc = parseXml('<w:t v="a&#9;b"/>');
    expect(serializeXml(doc)).toContain('&#x9;');
    expect(attr(parseXml(serializeXml(doc)).root, 'v')).toBe('a\tb');
  });

  it('空元素写成自闭合', () => {
    expect(serializeXml(parseXml('<w:br></w:br>'))).toBe('<w:br/>');
  });

  it('带回声明', () => {
    expect(serializeXml(parseXml(`${DECL}<w:p/>`))).toBe(`${DECL}<w:p/>`);
  });
});
