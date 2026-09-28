// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { collectFontFaces } from './snapshot-dom.ts';

describe('collectFontFaces', () => {
  // jsdom 的 CSSOM 会丢掉 src 描述符，规则原文与内联在浏览器回归（/tests/print.html ⑥）里验
  it('收顶层与 @media 里的 @font-face，族名去引号小写，相对地址的基准是文档地址', () => {
    const style = document.createElement('style');
    style.textContent = [
      '@font-face { font-family: "FangSong GB"; src: url(fonts/fs.woff2); }',
      'p { color: red; }',
      '@media screen { @font-face { font-family: KaiTi; src: url(kt.woff2); } }',
    ].join('\n');
    document.head.append(style);
    const faces = collectFontFaces(document);
    style.remove();
    expect(faces.map((f) => f.family)).toEqual(['fangsong gb', 'kaiti']);
    expect(faces[0]?.baseUrl).toBe(document.baseURI);
  });
});
