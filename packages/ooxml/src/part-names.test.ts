import { describe, expect, it } from 'vitest';
import { dirNameOf, extensionOf, relsPartNameOf, resolveTarget, toPartName } from './part-names.ts';

describe('resolveTarget', () => {
  it('相对路径基于源部件所在目录，不是 .rels 所在目录', () => {
    expect(resolveTarget('/word/document.xml', 'styles.xml')).toBe('/word/styles.xml');
    expect(resolveTarget('/word/document.xml', 'theme/theme1.xml')).toBe('/word/theme/theme1.xml');
  });

  it('折叠 ..', () => {
    expect(resolveTarget('/word/document.xml', '../docProps/app.xml')).toBe('/docProps/app.xml');
    expect(resolveTarget('/word/theme/theme1.xml', '../media/img1.png')).toBe('/word/media/img1.png');
  });

  it('绝对 Target 原样保留', () => {
    expect(resolveTarget('/word/document.xml', '/word/styles.xml')).toBe('/word/styles.xml');
  });

  it('包级关系（源为空串）落在根目录', () => {
    expect(resolveTarget('', 'word/document.xml')).toBe('/word/document.xml');
  });
});

describe('relsPartNameOf', () => {
  it('部件 → 同级 _rels 下的同名 .rels', () => {
    expect(relsPartNameOf('/word/document.xml')).toBe('/word/_rels/document.xml.rels');
  });

  it('包级关系是 /_rels/.rels', () => {
    expect(relsPartNameOf('')).toBe('/_rels/.rels');
    expect(relsPartNameOf('/')).toBe('/_rels/.rels');
  });
});

describe('杂项', () => {
  it('zip 条目名 → 部件名补前导斜杠', () => {
    expect(toPartName('word/document.xml')).toBe('/word/document.xml');
    expect(toPartName('/word/document.xml')).toBe('/word/document.xml');
  });

  it('dirNameOf 根部件返回空串', () => {
    expect(dirNameOf('/document.xml')).toBe('');
    expect(dirNameOf('/word/document.xml')).toBe('/word');
  });

  it('扩展名小写；目录里的点不算扩展名', () => {
    expect(extensionOf('/word/media/IMG1.PNG')).toBe('png');
    expect(extensionOf('/a.b/c')).toBe('');
  });
});
