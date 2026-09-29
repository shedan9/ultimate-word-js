/**
 * 环绕：方形 / 上下型环绕的浮动对象把文字**让开**。
 *
 * 浮动对象的位置与大小早就是对的（`page.ts` 的 `placeFloats`，参照框全部实测），
 * 这里只做「文字怎么让开」那一半。原来这一半没做，四周型的图直接压在文字上。
 *
 * 模型是**禁区**：每个环绕对象在本页上占一个矩形（外框按 `wp:anchor` 的 `dist*` 往外扩），
 * 一行文字的行盒 `[y, y + 行高)` 与禁区纵向相交时，这一行只能用禁区以外的那一段：
 *
 * - **上下型**（`topAndBottom`）：整行都不许放 —— 推到禁区底下
 * - **方形**（`square`；紧密型 / 穿越型按外接矩形退化成它，多边形绕排是写死的非目标）：
 *   行被禁区切成左右两段，按 `wrapText` 挑一段；挑出来的窄过 `WRAP_MIN_SEGMENT_EM` 个字
 *   就当放不下，同样推到禁区底下
 *
 * 推下去的那一截记在 `LineLayout.skip` 上（在行盒外面，不改基线）。
 *
 * 行高要等一行断完才知道，而这一行能用多宽又要看它的行盒碰不碰禁区 —— 所以拿**上一行的行高**
 * 当估计（首行拿不绕排时的首行）。同一段里行高几乎都一样，估错只在字号突变的那一行上。
 *
 * ⚠️ 没有真值（本机没有 Word）。几何上的判断全是按规范与界面行为推的：
 * 禁区按外框 + dist、行盒与禁区相交就算碰（碰边不算）、段落自己的缩进与禁区取交集、
 * 两侧都能放时 bothSides 退化成 largest（`WRAP_BOTH_SIDES_AS`）、最窄放一个字（`WRAP_MIN_SEGMENT_EM`）。
 * 钉死办法见 `uncalibrated.ts` 那两条。
 */
import type { Twips } from '@uw/core';
import type { DrawingAnchor } from '@uw/model';
import type { LineSlot } from './paragraph.ts';
import { WRAP_BOTH_SIDES_AS, WRAP_MIN_SEGMENT_EM } from './uncalibrated.ts';

/** 本页上一块不许放字的地方。坐标与分页的游标同一套：x 相对版心左边、y 相对版心顶 */
export interface WrapExclusion {
  left: Twips;
  right: Twips;
  top: Twips;
  bottom: Twips;
  mode: 'square' | 'topAndBottom';
  side: NonNullable<DrawingAnchor['wrapText']>;
}

/**
 * 一个浮动对象在本页上的禁区。`rect` 是它的纸坐标外框（`placeFloats` 算出来的那一份），
 * `origin` 是版心左上角的纸坐标。不让开文字的（`none`：衬于文字下方 / 浮于文字上方）答 undefined
 */
export function wrapExclusion(
  anchor: DrawingAnchor,
  rect: { x: Twips; y: Twips; width: Twips; height: Twips },
  origin: { x: Twips; y: Twips },
): WrapExclusion | undefined {
  if (anchor.wrap === 'none') return undefined;
  const d = anchor.dist;
  return {
    left: rect.x - d.left - origin.x,
    right: rect.x + rect.width + d.right - origin.x,
    top: rect.y - d.top - origin.y,
    bottom: rect.y + rect.height + d.bottom - origin.y,
    mode: anchor.wrap === 'topAndBottom' ? 'topAndBottom' : 'square',
    side: anchor.wrapText ?? 'bothSides',
  };
}

/** `slotAround` 的答案：一行摆在哪，外加「这一行是不是把 bothSides 退化了」（分页据此记诊断） */
export interface WrappedSlot extends LineSlot {
  bothSidesApproximated?: true;
}

/**
 * 一行摆在哪：从 `top`（这一行本来的行顶，版心坐标）开始往下找，找到第一个放得下的 y。
 *
 * @param base 不绕排时这一行的那一段（缩进已经算进去）
 * @param height 这一行的行高（估计值，见文件头）
 * @param unit 一个字宽
 */
export function slotAround(
  base: LineSlot,
  exclusions: readonly WrapExclusion[],
  top: Twips,
  height: Twips,
  unit: Twips,
): WrappedSlot {
  const lineL = base.left;
  const lineR = base.left + base.avail;
  const min = unit * WRAP_MIN_SEGMENT_EM;
  let y = top;
  // 每推一次至少越过一个禁区的底边，禁区数就是推的次数上限
  for (let guard = 0; guard <= exclusions.length; guard++) {
    const hits = exclusions.filter(
      (e) =>
        e.top < y + height &&
        e.bottom > y &&
        (e.mode === 'topAndBottom' || (e.right > lineL && e.left < lineR)),
    );
    if (hits.length === 0) return { ...base, skip: base.skip + y - top };

    if (hits.every((e) => e.mode === 'square')) {
      const segs = segmentsOf(lineL, lineR, hits).filter((s) => s.right - s.left >= min);
      const pick = choose(segs, hits);
      if (pick !== undefined) {
        const out: WrappedSlot = {
          left: pick.left,
          avail: pick.right - pick.left,
          skip: base.skip + y - top,
        };
        if (pick.approximated) out.bothSidesApproximated = true;
        return out;
      }
    }
    // 放不下：推到最先结束的那个禁区底下再试
    y = Math.min(...hits.map((e) => e.bottom));
  }
  return { ...base, skip: base.skip + y - top };
}

interface Segment {
  left: Twips;
  right: Twips;
}

/** `[lineL, lineR)` 减掉各禁区的横向投影，剩下的空当（按从左到右） */
function segmentsOf(lineL: Twips, lineR: Twips, hits: readonly WrapExclusion[]): Segment[] {
  const blocks = hits.map((e) => ({ left: Math.max(lineL, e.left), right: Math.min(lineR, e.right) }));
  blocks.sort((a, b) => a.left - b.left);
  const out: Segment[] = [];
  let x = lineL;
  for (const b of blocks) {
    if (b.left > x) out.push({ left: x, right: b.left });
    x = Math.max(x, b.right);
  }
  if (x < lineR) out.push({ left: x, right: lineR });
  return out;
}

/**
 * 按 `wrapText` 挑一段。`left` / `right` 说的是「文字只走对象的左 / 右边」，
 * 所以一个 `left` 的禁区把它右边的空当全否掉。剩下的空当里挑最宽的 ——
 * `largest` 本来就是这个意思，`bothSides` 两侧都能放时退化成它（见 `WRAP_BOTH_SIDES_AS`）
 */
function choose(
  segs: readonly Segment[],
  hits: readonly WrapExclusion[],
): (Segment & { approximated: boolean }) | undefined {
  const allowed = segs.filter((s) =>
    hits.every((e) =>
      e.side === 'left' ? s.right <= e.left : e.side === 'right' ? s.left >= e.right : true,
    ),
  );
  if (allowed.length === 0) return undefined;
  const widest = allowed.reduce((a, b) => (b.right - b.left > a.right - a.left ? b : a));
  const approximated =
    WRAP_BOTH_SIDES_AS === 'largest' &&
    allowed.length > 1 &&
    hits.some(
      (e) =>
        e.side === 'bothSides' &&
        allowed.some((s) => s.right <= e.left) &&
        allowed.some((s) => s.left >= e.right),
    );
  return { ...widest, approximated };
}
