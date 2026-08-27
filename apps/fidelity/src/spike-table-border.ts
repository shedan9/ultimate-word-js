#!/usr/bin/env node
/**
 * 格线冲突穿刺：一条格线两边各写了一条**不一样**的边框时，Word 画的是哪一条。
 *
 * 这一级此前完全没有真值 —— `layout/src/table-borders.ts` 的「相邻竞争」照 CSS 2.1
 * §17.6.2 的 collapsing borders 类比写成「先比线宽、再比样式权重、仍平局取左上」。
 * 样本一跑就照出两处错：**破折类（虚线 / 点线）再宽也输给实线**，
 * 而**同一种破折线之间连宽度都不比**。
 *
 * ## 样本怎么造的
 *
 * `spike-table-03`：21 组配对，每组做两遍（竖边一行两格 / 横边两行一格）。
 * 竞争的两侧各给一个独一无二的**颜色**（左 / 上 = 红 FF0000，右 / 下 = 蓝 0000FF，
 * 退到表级的 `insideH` / `insideV` = 绿 008000），于是「画出来的是哪一条」直接从
 * 真值里那条线的颜色读出来 —— 与 `spike-table-02` 拿字号认条件格式同一招。
 * 冲突的 `w:tcBorders` 是**改 XML** 写进去的（`patch-docx.ts`）：Word 的对象模型里
 * 一条共享边只有一个 Border 对象，设一侧等于两侧都设，它自己造不出这个局面。
 *
 * ## 判据
 *
 * 与别的 spike 一样是**排组合**（`BorderConflictRules` 的 2×2×2×2×2 = 32 种），
 * 逐条边比「画出来的颜色」，唯一满分的那一组就是实现该用的那一组。
 * 残差在这里没有意义：赢家是离散的，一条边只有对与不对。
 *
 * 这个脚本**不需要 Word**：docx 与 truth.json 都入库了。
 * 重新造样本才要 Windows（`pnpm truth spike-table-03 --force`）。
 *
 *   node src/spike-table-border.ts
 */
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createDiagnosticSink } from '@uw/core';
import type { BorderConflictRules } from '@uw/layout';
import { BORDER_CONFLICT_RULES, borderRowsOf, resolveTableBorders } from '@uw/layout';
import type { ResolvedBlock, ResolvedTable } from '@uw/model';
import { loadDocument, paragraphText } from '@uw/model';
import { OpcPackage } from '@uw/ooxml';
import type { TruthPage, TruthRule, WordTruth } from './truth-types.ts';

const APP_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const FIXTURE = 'spike-table-03';

/** 外框与交点方块都是黑的；竞争的读数只看这三种颜色 */
const MARKED = new Set(['#ff0000', '#0000ff', '#008000']);

/** 一条竞争边：从样本里读出来的「谁赢了」 */
interface Site {
  /** 组名（〇 一 二 …）+ 方向 */
  name: string;
  /** Word 画出来的那条线是谁的：红 / 蓝 / 绿 / 无 */
  observed: string;
  /** 我们解析出来的边框颜色（`undefined` = 不画） */
  predicted: (rules: BorderConflictRules) => string;
}

const colorName = (c: string | undefined): string =>
  c === undefined
    ? '无'
    : c.toLowerCase() === 'ff0000'
      ? '红'
      : c.toLowerCase() === '0000ff'
        ? '蓝'
        : c.toLowerCase() === '008000'
          ? '绿'
          : c;

const observedName = (rules: readonly TruthRule[]): string => {
  const marked = rules.filter((r) => MARKED.has(r.color.toLowerCase()));
  if (marked.length === 0) return '无';
  const first = marked[0] as TruthRule;
  return colorName(first.color.slice(1));
};

// ── 真值一侧：每组的那条线是什么颜色 ──────────────────────────────────────────
//
// 定位不靠坐标算，靠**格子里的文字**：每组的两格写着「一左 / 一右」或「一上 / 一下」，
// 真值里同一基线的片段已经拼成一行，于是「这一组的表在哪一段 y 上」直接读得到。
// 颜色只在竞争的那条边上出现（外框与交点方块全是黑的），所以那一段里唯一的彩色矩形
// 就是答案 —— 不需要知道列边界在哪，也就不会把已经标定完的几何牵进这份结论。

interface Observed {
  vertical: Map<string, string>;
  horizontal: Map<string, string>;
}

function observe(truth: WordTruth): Observed {
  const vertical = new Map<string, string>();
  const horizontal = new Map<string, string>();
  for (const page of truth.pages) {
    for (const line of page.lines) {
      const text = line.text.replace(/\s+/gu, '');
      const v = /^(.)左(.)右$/u.exec(text);
      if (v !== null && v[1] !== undefined) {
        // 竖边横跨这一行：取纵向盖住基线上方那一格的矩形
        vertical.set(v[1], observedName(inBand(page, line.y - 12, line.y + 2, 'vertical')));
      }
      const h = /^(.)上$/u.exec(text);
      if (h !== null && h[1] !== undefined) {
        const below = page.lines.find((o) => o.text.replace(/\s+/gu, '') === `${h[1]}下`);
        if (below === undefined) continue;
        // 横边落在两行文字之间；下一行的基线上方留 6pt 给那一行自己的格线
        horizontal.set(h[1], observedName(inBand(page, line.y + 1, below.y - 6, 'horizontal')));
      }
    }
  }
  return { vertical, horizontal };
}

/** 一段 y 区间里的线。`vertical` 要求纵向盖住整段（竖线），`horizontal` 要求线心落在段内 */
function inBand(page: TruthPage, y0: number, y1: number, kind: 'vertical' | 'horizontal'): TruthRule[] {
  return (page.rules ?? []).filter((r) =>
    kind === 'vertical' ? r.y < y1 && r.y + r.h > y0 && r.h > 5 : r.y >= y0 && r.y <= y1 && r.w > 5,
  );
}

// ── 我们这一侧：同一组配对解析出来的是哪条边 ──────────────────────────────────

function* tables(blocks: readonly ResolvedBlock[]): Generator<ResolvedTable> {
  for (const b of blocks) {
    if (b.kind !== 'table') continue;
    yield b;
  }
}

/** 表格第一格的文字，用来认出这是哪一组（「一左」/「一上」） */
function firstCellText(table: ResolvedTable): string {
  const cell = table.rows[0]?.cells[0];
  if (cell === undefined) return '';
  return cell.blocks
    .map((b) => (b.kind === 'paragraph' ? paragraphText(b) : ''))
    .join('')
    .replace(/\s+/gu, '');
}

function sitesOf(doc: ReturnType<typeof loadDocument>, observed: Observed): Site[] {
  const out: Site[] = [];
  for (const section of doc.resolved.sections) {
    for (const table of tables(section.blocks)) {
      const label = firstCellText(table)[0];
      if (label === undefined) continue;
      const rows = borderRowsOf(table.rows);
      const colCount = table.grid.length;
      const vertical = table.rows.length === 1;
      const want = (vertical ? observed.vertical : observed.horizontal).get(label);
      if (want === undefined) continue;
      out.push({
        name: `${label}${vertical ? '竖' : '横'}`,
        observed: want,
        predicted: (rules) => {
          const solved = resolveTableBorders(rows, table.props.borders, colCount, rules);
          const border = vertical
            ? solved[0]?.[0]?.right
            : // 横边：上面那格的 bottom 与下面那格的 top 解析出的是同一条，取前者
              solved[0]?.[0]?.bottom?.[0]?.border;
          return colorName(border?.color);
        },
      });
    }
  }
  return out;
}

// ── 排组合 ────────────────────────────────────────────────────────────────────

const CANDIDATES = {
  order: ['class', 'thickness'],
  brokenByThickness: [false, true],
  thickness: ['rendered', 'size'],
  styleBreaksTie: [true, false],
  position: ['leftTop', 'rightBottom'],
} as const satisfies { [K in keyof BorderConflictRules]: readonly BorderConflictRules[K][] };

const combos: BorderConflictRules[] = [];
for (const order of CANDIDATES.order) {
  for (const brokenByThickness of CANDIDATES.brokenByThickness) {
    for (const thickness of CANDIDATES.thickness) {
      for (const styleBreaksTie of CANDIDATES.styleBreaksTie) {
        for (const position of CANDIDATES.position) {
          combos.push({ order, brokenByThickness, thickness, styleBreaksTie, position });
        }
      }
    }
  }
}

const label = (r: BorderConflictRules): string =>
  `${r.order === 'class' ? '先分类' : '只比厚'} / 破折${r.brokenByThickness ? '比宽' : '不比宽'}` +
  ` / 厚度${r.thickness === 'rendered' ? '按画出来' : '按 sz'}` +
  ` / ${r.styleBreaksTie ? '样式破平局' : '样式不参与'} / 平局取${r.position === 'leftTop' ? '左上' : '右下'}`;

const isDefault = (r: BorderConflictRules): boolean =>
  (Object.keys(CANDIDATES) as (keyof BorderConflictRules)[]).every((k) => r[k] === BORDER_CONFLICT_RULES[k]);

const dir = path.join(APP_ROOT, 'fixtures');
const bytes = new Uint8Array(await readFile(path.join(dir, `${FIXTURE}.docx`)));
const truth = JSON.parse(await readFile(path.join(dir, `${FIXTURE}.truth.json`), 'utf8')) as WordTruth;
const doc = loadDocument(OpcPackage.open(bytes), createDiagnosticSink());
const sites = sitesOf(doc, observe(truth));

console.log('格线冲突穿刺 —— 相邻两格各写一条边，Word 画哪一条（排组合）\n');
if (sites.length === 0) throw new Error('一组都没配上 —— 样本或定位规则变了');

const scored = combos.map((rules) => ({
  rules,
  hit: sites.filter((s) => s.predicted(rules) === s.observed).length,
}));
for (const s of scored) {
  console.log(`${isDefault(s.rules) ? '→ ' : '  '}${label(s.rules).padEnd(58)}${s.hit} / ${sites.length}`);
}

const best = Math.max(...scored.map((s) => s.hit));
const winners = scored.filter((s) => s.hit === best);
console.log(`\n最优：${best} / ${sites.length}，共 ${winners.length} 种组合并列`);
for (const w of winners) console.log(`  · ${label(w.rules)}`);

console.log('\n实现的这一组，逐组配对：');
for (const s of sites) {
  const got = s.predicted(BORDER_CONFLICT_RULES);
  console.log(`  ${s.name.padEnd(4)} Word ${s.observed}  /  我们 ${got}${got === s.observed ? '' : '   ✗'}`);
}

const only = winners[0];
const ok = best === sites.length && winners.length === 1 && only !== undefined && isDefault(only.rules);
console.log(
  ok
    ? `\n✓ 「先分类 / 破折不比宽 / 厚度按画出来算 / 样式破平局 / 平局取左上」是唯一满分的一组（${sites.length} 条边全对）`
    : '\n✗ 与真值对不上 —— 见上表',
);
if (!ok) process.exitCode = 1;
