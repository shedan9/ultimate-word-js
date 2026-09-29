/**
 * 脚注跨页续排 —— 把排好的一条脚注（`StackResult`）按**行**切成几截。
 *
 * 脚注内容与页眉一样先整条摞好（`stackBlocks`，缓存按「哪一条 × 哪一节」），
 * 分页只决定「这一页收到哪一行为止」。所以一截就是摞好的坐标里的一个区间 `[top, bottom)`，
 * 不重新排版：续到下一页的那几行与留在本页的那几行出自**同一份**排版结果，
 * 断行点、行高一个都不会因为被切开而变。
 *
 * 切口只落在**行与行之间**（段落的行、表格的整行）：半行字没有意义，
 * 表格行内部再切要走 `table-split.ts` 那一套，脚注里的表格罕见到不值得。
 * 两截之间的那段空白（段后 + 下一段段前）两边都不要：本页那一截到切口那一行的行底为止，
 * 续页那一截从下一行的行顶开始 —— 与正文「段前间距落在页首不算」同一个道理。
 *
 * 纯函数，不认识分页；分页那一侧的规则（引用页至少留几行、续页顶上画什么线）在 `page.ts`。
 */
import type { Twips } from '@uw/core';
import type { StackResult } from './header-footer.ts';
import type { PlacedBlock } from './page.ts';

/**
 * 一截脚注。`top` / `bottom` 是这条脚注**摞好时自己的坐标**，高度 = `bottom - top`。
 * `section` 是按哪一节的版心宽排的 —— 续到下一节的页上时仍用原来那一份，
 * 否则两截的切口坐标对不上（两节版心宽不同，行数都可能不同）。
 */
export interface NotePiece {
  key: string;
  section: number;
  top: Twips;
  bottom: Twips;
}

export interface PieceFit {
  /** 这一截的底（摞好的坐标） */
  bottom: Twips;
  /** 收进来几行（段落的行 / 表格的整行） */
  lines: number;
  /** 剩下的全收完了 */
  complete: boolean;
}

interface Unit {
  top: Twips;
  bottom: Twips;
}

/** 可以切开的最小单位：段落的每一行、表格的每一整行，按从上到下 */
function unitsOf(stacked: StackResult): Unit[] {
  const out: Unit[] = [];
  for (const b of stacked.blocks) {
    if (b.kind === 'paragraph') for (const l of b.lines) out.push({ top: l.y, bottom: l.y + l.line.height });
    else for (const r of b.rows) out.push({ top: r.y, bottom: r.y + r.height });
  }
  return out;
}

/**
 * 从 `top` 起、`room` 这么高的地方最多收到哪一行。
 *
 * 整条剩下的都放得下时底取 `stacked.height`（带着末尾的段后间距，与不切开时让出的高一致）；
 * 行都放得下、只差末尾那段段后间距时也算收完，底取末行行底 —— 切口在页底，那段空白画不画都一样，
 * 为它把最后一行挪到下一页只会凭空多出一截续排。
 * `force` 时至少收一行（一行脚注比整页版心还高，只能溢出，否则续排永远续不完）。
 */
export function fitPiece(stacked: StackResult, top: Twips, room: Twips, force = false): PieceFit {
  const units = unitsOf(stacked).filter((u) => u.top >= top);
  if (stacked.height - top <= room) return { bottom: stacked.height, lines: units.length, complete: true };
  let bottom = top;
  let lines = 0;
  for (const u of units) {
    if (u.bottom - top > room && !(force && lines === 0)) break;
    bottom = u.bottom;
    lines += 1;
  }
  return { bottom, lines, complete: lines === units.length };
}

/** 切口 `bottom` 之后，下一截从哪儿开始（下一行的行顶）；已经收完时是 undefined */
export function nextPieceTop(stacked: StackResult, bottom: Twips): Twips | undefined {
  for (const u of unitsOf(stacked)) if (u.top >= bottom) return u.top;
  return undefined;
}

/**
 * 取出 `[top, bottom)` 那一截的块，坐标挪到以这一截的顶为 0。
 * 缓存里的块是共享的，这里只造新对象、不就地改。
 */
export function slicePiece(stacked: StackResult, top: Twips, bottom: Twips): PlacedBlock[] {
  const out: PlacedBlock[] = [];
  const inside = (y: Twips, h: Twips) => y >= top && y + h <= bottom;
  for (const b of stacked.blocks) {
    if (b.kind === 'paragraph') {
      const lines = b.lines.filter((l) => inside(l.y, l.line.height));
      const first = lines[0];
      if (first === undefined) continue;
      // 整块都在这一截里就原样挪过去：`y` 不含段前间距，按首行重算会丢掉整块时的那份
      if (lines.length === b.lines.length) {
        out.push({ ...b, y: b.y - top, lines: lines.map((l) => ({ ...l, y: l.y - top })) });
        continue;
      }
      out.push({
        ...b,
        y: first.y - top,
        lines: lines.map((l) => ({ ...l, y: l.y - top })),
        first: b.first && first === b.lines[0],
        last: b.last && lines.at(-1) === b.lines.at(-1),
      });
    } else {
      const rows = b.rows.filter((r) => inside(r.y, r.height));
      const first = rows[0];
      if (first === undefined) continue;
      out.push({
        ...b,
        y: (rows.length === b.rows.length ? b.y : first.y) - top,
        rows: rows.map((r) => ({ ...r, y: r.y - top })),
        first: b.first && first === b.rows[0],
        last: b.last && rows.at(-1) === b.rows.at(-1),
      });
    }
  }
  return out;
}
