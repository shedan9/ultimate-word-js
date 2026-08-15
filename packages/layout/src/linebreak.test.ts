/**
 * 断行：禁则、挤压、悬挂、回退、制表位、硬换行。
 *
 * 判据是「断在第几个字」，所以合成字体取了最好算的形状：一个汉字 = 一个字号宽，
 * 可用宽度写成 `SIZE_5 * n` 就是「这一行放得下 n 个字」。
 */
import { describe, expect, it } from 'vitest';
import { buildItems } from './items.ts';
import type { LineBreakContext } from './linebreak.ts';
import { breakLines } from './linebreak.ts';
import { fakeMeasurer, para, run, runOf, SIZE_5 } from './test-fixtures.ts';
import type { LayoutItem } from './types.ts';

const M = { measurer: fakeMeasurer() };

function ctx(chars: number, over: Partial<LineBreakContext> = {}): LineBreakContext {
  return {
    availWidth: () => SIZE_5 * chars,
    lineLeft: () => 0,
    tabs: [],
    defaultTabStop: 420,
    compressPunctuation: false,
    overflowPunct: false,
    ...over,
  };
}

/** 把断行结果还原成每行的文字，肉眼可读的断言比下标好使 */
function texts(items: readonly LayoutItem[], lines: { start: number; end: number }[]): string[] {
  return lines.map((l) =>
    items
      .slice(l.start, l.end)
      .map((i) => (i.kind === 'char' ? String.fromCodePoint(i.cp) : i.kind === 'tab' ? '\t' : ''))
      .join(''),
  );
}

describe('基本断行', () => {
  it('中文按可用宽度逐字断 —— 汉字之间处处可断', () => {
    const items = buildItems(para([run('一二三四五六七八九十')]), M);
    expect(texts(items, breakLines(items, ctx(4)))).toEqual(['一二三四', '五六七八', '九十']);
  });

  it('拉丁词内不断，回退到词首（空格之后才是断点）', () => {
    const items = buildItems(para([run('hello world')]), M);
    // 半角 0.5 em：8 个字号宽 = 16 个半角字符，'hello world' 共 11 个 → 一行放得下
    expect(texts(items, breakLines(items, ctx(4)))).toEqual(['hello ', 'world']);
  });

  it('一个 item 比整行还宽也得收下，不能死循环', () => {
    const items = buildItems(para([run('中中')]), M);
    expect(texts(items, breakLines(items, ctx(0.5)))).toEqual(['中', '中']);
  });

  it('空段落也产出一行 —— 空行同样占高度', () => {
    expect(breakLines([], ctx(10))).toHaveLength(1);
  });
});

describe('禁则（避头尾）', () => {
  it('后置标点不许出现在行首：回退，把前一个字一起推下去', () => {
    const items = buildItems(para([run('一二三四。五')]), M);
    // 4 个字宽的行放得下「一二三四」，但「。」不能起头 → 回退到「四。」一起下行
    expect(texts(items, breakLines(items, ctx(4)))).toEqual(['一二三', '四。五']);
  });

  it('前置标点不许留在行尾', () => {
    const items = buildItems(para([run('一二三（四）')]), M);
    expect(texts(items, breakLines(items, ctx(4)))).toEqual(['一二三', '（四）']);
  });

  it('整行没有合法断点时硬断，不至于把一行撑到无限长', () => {
    const items = buildItems(para([run('。。。。。。')]), M);
    expect(texts(items, breakLines(items, ctx(2)))).toEqual(['。。', '。。', '。。']);
  });
});

describe('压缩优先，压不下再回退', () => {
  it('挤压：行尾全角标点压掉空着的半边就塞得下', () => {
    const items = buildItems(para([run('一二。三')]), M);
    const lines = breakLines(items, ctx(2.5, { compressPunctuation: true }));
    expect(texts(items, lines)[0]).toBe('一二。');
    // 压过之后这一行正好占满：两个全角字 + 半个标点
    expect(lines[0]?.width).toBe(SIZE_5 * 2.5);
  });

  it('悬挂：overflowPunct 开着时行尾标点吐出版心，且不计入行宽', () => {
    const items = buildItems(para([run('一二。三')]), M);
    const lines = breakLines(items, ctx(2, { overflowPunct: true }));
    expect(texts(items, lines)[0]).toBe('一二。');
    expect(lines[0]?.width).toBe(SIZE_5 * 2);
    expect(lines[0]?.hanging).toEqual([false, false, true]);
  });

  it('两条补救都关掉才走回退 —— 这三条的顺序决定断在第几个字', () => {
    const items = buildItems(para([run('一二。三')]), M);
    // 「。」既压不了也不许悬挂，只能连着前一个字下行；下一行它又正好排得下
    expect(texts(items, breakLines(items, ctx(2)))).toEqual(['一', '二。', '三']);
  });

  it('行尾空格允许溢出，不计入行宽 —— 否则末尾多打一个空格就会让居中标题左移', () => {
    const items = buildItems(para([run('一二 三')]), M);
    const lines = breakLines(items, ctx(2));
    expect(lines[0]?.width).toBe(SIZE_5 * 2);
  });
});

describe('制表位', () => {
  const tabbed = () => buildItems(para([runOf([{ kind: 'tab' }, { kind: 'text', text: '中' }])]), M);

  it('没有显式制表位时推进到 defaultTabStop 的整数倍', () => {
    const items = tabbed();
    const line = breakLines(items, ctx(10))[0];
    expect(line?.ws[0]).toBe(420);
    expect(line?.xs[1]).toBe(420);
  });

  it('显式制表位优先，bar 型不参与推进（它只是一条竖线）', () => {
    const items = tabbed();
    const line = breakLines(
      items,
      ctx(10, {
        tabs: [
          { pos: 300, alignment: 'bar', leader: 'none' },
          { pos: 1000, alignment: 'left', leader: 'dot' },
        ],
      }),
    )[0];
    expect(line?.ws[0]).toBe(1000);
    expect(line?.tabs[0]).toMatchObject({ alignment: 'left', leader: 'dot', pos: 1000 });
  });

  it('右对齐制表位把后面那段拉到停靠点上结束', () => {
    const items = buildItems(
      para([runOf([{ kind: 'text', text: '一' }, { kind: 'tab' }, { kind: 'text', text: '二三' }])]),
      M,
    );
    const line = breakLines(items, ctx(20, { tabs: [{ pos: 2000, alignment: 'right', leader: 'none' }] }))[0];
    // 「二三」两个全角字，末尾对齐到 2000
    expect((line?.xs[3] as number) + SIZE_5).toBe(2000);
  });

  it('制表位按版心的绝对坐标算，不是按行首 —— 缩进过的行同样对得齐', () => {
    const items = tabbed();
    const line = breakLines(items, ctx(10, { lineLeft: () => 500 }))[0];
    expect(line?.ws[0]).toBe(420 * 2 - 500); // 500 之后的下一个 420 整数倍是 840
  });
});

describe('硬换行', () => {
  it('w:br 结束当前行，并把类型带出去给分页用', () => {
    const items = buildItems(
      para([
        runOf([
          { kind: 'text', text: '一' },
          { kind: 'break', breakType: 'page' },
          { kind: 'text', text: '二' },
        ]),
      ]),
      M,
    );
    const lines = breakLines(items, ctx(10));
    expect(lines).toHaveLength(2);
    expect(lines[0]?.breakAfter).toBe('page');
  });
});

describe('结构化克隆', () => {
  it('断行结果是纯数据 —— 它要过 Worker 边界', () => {
    const items = buildItems(para([run('一二三四五')]), M);
    const lines = breakLines(items, ctx(2));
    expect(structuredClone(lines)).toEqual(lines);
    expect(structuredClone(items)).toEqual(items);
  });
});
