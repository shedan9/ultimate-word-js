import type { LineFragment, LineObject, PageLayout } from '@uw/layout';
import { describe, expect, it } from 'vitest';
import type { FetchBytes, FontFaceSource } from './snapshot.ts';
import { dataUri, familiesOf, pageSnapshotMarkup } from './snapshot.ts';

const STYLE: LineFragment['style'] = {
  bold: false,
  italic: false,
  color: 'auto',
  underline: 'none',
  strike: false,
  doubleStrike: false,
  vertAlign: 'baseline',
  position: 0,
  scale: 100,
};

function fragment(font: string): LineFragment {
  return {
    runId: 'r1',
    contentIndex: 0,
    offset: 0,
    font,
    fontSize: 320,
    script: 'eastAsia',
    text: '公文',
    x: 0,
    width: 640,
    glyphX: [0, 320],
    style: STYLE,
  };
}

function object(id: string): LineObject {
  return {
    runId: 'r2',
    contentIndex: 0,
    x: 640,
    width: 1440,
    height: 1440,
    objectKind: 'drawing',
    image: { id, relId: id },
  };
}

function page(font = '仿宋', objects: LineObject[] = []): PageLayout {
  return {
    index: 0,
    number: 1,
    sectionIndex: 0,
    geometry: { width: 11906, height: 16838, content: { x: 1800, y: 1440, width: 8306, height: 13958 } },
    blocks: [
      {
        kind: 'paragraph',
        id: 'p1',
        y: 0,
        lines: [
          {
            index: 0,
            y: 0,
            line: {
              start: 0,
              end: 3,
              x: 0,
              width: 2080,
              height: 1440,
              baseline: 1200,
              natural: 1440,
              fragments: [fragment(font)],
              leaders: [],
              objects,
              isLast: true,
            },
          },
        ],
        first: true,
        last: true,
      },
    ],
  };
}

const FONT_BYTES = new Uint8Array([0x77, 0x4f, 0x46, 0x32]);

/** 记下被取过哪些地址；`fail` 里的地址答 404 */
function fetcher(fail: string[] = []): FetchBytes & { calls: string[] } {
  const calls: string[] = [];
  const fn = async (url: string) => {
    calls.push(url);
    if (fail.some((f) => url.endsWith(f))) throw new Error('404');
    return { bytes: FONT_BYTES, type: url.endsWith('.png') ? 'image/png' : '' };
  };
  return Object.assign(fn, { calls });
}

const face = (family: string, src: string): FontFaceSource => ({
  family,
  cssText: `@font-face { font-family: "${family}"; src: ${src}; }`,
  baseUrl: 'https://host.example/css/app.css',
});

describe('familiesOf', () => {
  it('拆开字体栈、去引号、小写 —— 与 @font-face 的族名同一种写法才比得上', () => {
    expect(familiesOf('FangSong, \'仿宋\', "Noto Serif CJK SC", serif')).toEqual([
      'fangsong',
      '仿宋',
      'noto serif cjk sc',
      'serif',
    ]);
  });
});

describe('pageSnapshotMarkup', () => {
  it('是能独立加载的 SVG：带命名空间，宽高按 scale 出 px，viewBox 仍是 pt', async () => {
    const svg = await pageSnapshotMarkup(page(), { zoom: 2 });
    expect(svg.startsWith('<svg xmlns="http://www.w3.org/2000/svg"')).toBe(true);
    // A4 = 11906 twips = 793.73px @1x
    expect(svg).toContain('width="1587.467px"');
    expect(svg).toContain('viewBox="0 0 595.3 841.9"');
  });

  it('只内联这一页用到的字体族，地址相对样式表解析、没带类型时按扩展名补', async () => {
    const fetchBytes = fetcher();
    const svg = await pageSnapshotMarkup(page(), {
      fontFaces: [
        face('fangsong', "url('../fonts/fs.woff2') format('woff2')"),
        face('kaiti', 'url(kt.woff2)'),
      ],
      fetchBytes,
    });
    expect(fetchBytes.calls).toEqual(['https://host.example/fonts/fs.woff2']);
    expect(svg).toContain(
      `<style>@font-face { font-family: "fangsong"; src: url("${dataUri(FONT_BYTES, 'font/woff2')}")`,
    );
    expect(svg).not.toContain('kaiti');
    // style 必须在最前面：放在文字后面也生效，但序列化顺序固定下来快照才稳定
    expect(svg.indexOf('<style>')).toBeLessThan(svg.indexOf('<text'));
  });

  it('一个地址都取不到的 @font-face 整条丢掉，免得占着族名挡住本机同名字体', async () => {
    const svg = await pageSnapshotMarkup(page(), {
      fontFaces: [face('fangsong', 'url(fs.woff2)')],
      fetchBytes: fetcher(['fs.woff2']),
    });
    expect(svg).not.toContain('<style>');
  });

  it('src 里有一项取到就留下，取不到的那项保留原地址', async () => {
    const svg = await pageSnapshotMarkup(page(), {
      fontFaces: [face('fangsong', 'url(a.woff2), url(b.ttf)')],
      fetchBytes: fetcher(['a.woff2']),
    });
    expect(svg).toContain('url(a.woff2)');
    expect(svg).toContain(`url("${dataUri(FONT_BYTES, 'font/ttf')}")`);
  });

  it('非 data URI 的图片换成 data URI，data URI 的不重取', async () => {
    const fetchBytes = fetcher();
    const svg = await pageSnapshotMarkup(page('仿宋', [object('a'), object('b')]), {
      imageHref: (id) => (id === 'a' ? 'blob:https://host.example/1.png' : 'data:image/png;base64,AAAA'),
      fetchBytes,
    });
    expect(fetchBytes.calls).toEqual(['blob:https://host.example/1.png']);
    expect(svg).toContain(`href="${dataUri(FONT_BYTES, 'image/png')}"`);
    expect(svg).toContain('href="data:image/png;base64,AAAA"');
    expect(svg).not.toContain('blob:');
  });

  it('dataUri 分块编码，大于一块的字节也不爆调用栈', () => {
    const big = new Uint8Array(0x8000 * 3 + 5).fill(65);
    const uri = dataUri(big, 'font/woff2');
    expect(atob(uri.slice('data:font/woff2;base64,'.length))).toBe('A'.repeat(big.length));
  });
});
