/**
 * 格线**冲突**的真值回归：`spike-table-03`，21 组配对 × 横竖两遍，逐条边与 Word 对。
 *
 * 在这份样本之前，「共享一条线的两个格子各写了一条不同的边框，画哪一条」整条规则是照
 * CSS 2.1 §17.6.2 的 collapsing borders 类比出来的，一行都没跟 Word 比过。规则怎么标定
 * 出来的、每一组各回答什么，见 `apps/fidelity` 的 `spike:table-border` 与
 * `table-borders.ts` 的 `BORDER_CONFLICT_RULES` 证据表。
 *
 * 读数靠**颜色**：竞争的两侧各给一个独一无二的颜色（左 / 上 = 红 FF0000、
 * 右 / 下 = 蓝 0000FF、退到表级 = 绿 008000），Word 画出来的那条线是什么颜色，
 * 就是赢家是谁。所以这份回归**不碰几何**：它比的是颜色，不是坐标 ——
 * 断行或行高哪天变了也不该让它变红。
 *
 * 定位同样不靠坐标：每组两格写着「一左 / 一右」或「一上 / 一下」，真值里同基线的片段
 * 已经拼成一行，于是「这一组在哪一段 y 上」直接读得到，那一段里唯一的彩色矩形就是答案。
 *
 * 跨平台：docx 与 truth.json 都入库，这里只解析模型不排版，CI 上照跑。
 */
import { readFileSync } from 'node:fs';
import { createDiagnosticSink } from '@uw/core';
import type { ResolvedBlock, ResolvedTable } from '@uw/model';
import { loadDocument, paragraphText } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import { describe, expect, it } from 'vitest';
import { borderRowsOf, resolveTableBorders } from './table-borders.ts';

const FIXTURES = new URL('../../../apps/fidelity/fixtures/', import.meta.url);

interface Rule {
  x: number;
  y: number;
  w: number;
  h: number;
  color: string;
}
interface Truth {
  pages: { lines: { y: number; text: string }[]; rules?: Rule[] }[];
}

/** 外框与交点方块全是黑的，只有竞争的那条边带色 */
const MARKED: Record<string, string> = { '#ff0000': '红', '#0000ff': '蓝', '#008000': '绿' };

const colorName = (c: string | undefined): string =>
  c === undefined ? '无' : (MARKED[`#${c.toLowerCase()}`] ?? c);

const bytes = new Uint8Array(readFileSync(new URL('spike-table-03.docx', FIXTURES)));
const truth = JSON.parse(readFileSync(new URL('spike-table-03.truth.json', FIXTURES), 'utf8')) as Truth;
const doc = loadDocument(OpcPackage.open(bytes), createDiagnosticSink());

/** Word 画出来的是哪一条：组名 + 方向 → 颜色 */
function observed(): Map<string, string> {
  const out = new Map<string, string>();
  const pick = (rules: readonly Rule[]): string => {
    const first = rules.find((r) => MARKED[r.color.toLowerCase()] !== undefined);
    return first === undefined ? '无' : colorName(first.color.slice(1));
  };
  for (const page of truth.pages) {
    for (const line of page.lines) {
      const text = line.text.replace(/\s+/gu, '');
      const rules = page.rules ?? [];
      const v = /^(.)左(.)右$/u.exec(text);
      if (v?.[1] !== undefined) {
        // 竖线纵向盖住这一行
        out.set(
          `${v[1]}竖`,
          pick(rules.filter((r) => r.h > 5 && r.y < line.y + 2 && r.y + r.h > line.y - 12)),
        );
      }
      const h = /^(.)上$/u.exec(text);
      if (h?.[1] !== undefined) {
        const below = page.lines.find((o) => o.text.replace(/\s+/gu, '') === `${h[1]}下`);
        if (below === undefined) continue;
        // 横线落在两行文字之间，下一行基线上方留 6pt 给那一行自己的格线
        out.set(`${h[1]}横`, pick(rules.filter((r) => r.w > 5 && r.y >= line.y + 1 && r.y <= below.y - 6)));
      }
    }
  }
  return out;
}

function* tables(blocks: readonly ResolvedBlock[]): Generator<ResolvedTable> {
  for (const b of blocks) {
    if (b.kind === 'table') yield b;
  }
}

/** 我们解析出来的是哪一条：组名 + 方向 → 颜色 */
function resolved(): Map<string, string> {
  const out = new Map<string, string>();
  for (const section of doc.resolved.sections) {
    for (const table of tables(section.blocks)) {
      const label = table.rows[0]?.cells[0]?.blocks
        .map((b) => (b.kind === 'paragraph' ? paragraphText(b) : ''))
        .join('')
        .replace(/\s+/gu, '')[0];
      if (label === undefined) continue;
      const solved = resolveTableBorders(borderRowsOf(table.rows), table.props.borders, table.grid.length);
      // 一行两格的表比竖边，两行一格的表比横边
      const vertical = table.rows.length === 1;
      const border = vertical ? solved[0]?.[0]?.right : solved[0]?.[0]?.bottom?.[0]?.border;
      out.set(`${label}${vertical ? '竖' : '横'}`, colorName(border?.color));
    }
  }
  return out;
}

describe('spike-table-03：格线冲突逐条边比真值', () => {
  const want = observed();
  const got = resolved();

  it('样本里的 42 条竞争边一条不少', () => {
    // 21 组 × 横竖两遍。少了说明样本或定位规则变了，而不是规则对不上
    expect(want.size).toBe(42);
    expect(got.size).toBe(42);
  });

  it('每一条边画出来的都是同一侧那条', () => {
    const wrong = [...want]
      .filter(([k, v]) => got.get(k) !== v)
      .map(([k, v]) => `${k}：Word ${v} / 我们 ${got.get(k)}`);
    expect(wrong).toEqual([]);
  });
});
