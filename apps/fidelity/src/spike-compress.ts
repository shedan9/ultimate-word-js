#!/usr/bin/env node
/**
 * 临时挤压穿刺：一行塞不下最后一个字时，Word 肯挤掉多少标点才肯换行。
 *
 * 这一条卡了 L2 很久：gongwen-01 的真值只给出一个区间（接受过 9.30pt 的挤压、拒绝过
 * 13.75pt），而且按标点算的模型在它上面自相矛盾。两份专门的样本把变量拆开了 ——
 * 每段都是**恰好一行放不下最后一个字**，靠右缩进（单位 pt）把可用宽度一格格调窄，
 * 于是「留住最后一个字要挤掉多少」这个量连续可控。
 *
 * 三条结论：
 *
 * 1. **左对齐一格都不挤**（`spike-compress-01` 的 B 组，15 段全部换行）。
 *    挤压是两端对齐才有的行为 —— 左对齐的右边本来就是毛边，挤出来的地方没有用处
 * 2. 一个标点最多让出 **0.48 em**（`spike-compress-02` 的 G1 组）
 * 3. 挤到多少就宁可换行，是拿它跟**换行后要拉开的量**比出来的：
 *    `挤压量 × 字距数 ≤ K × 标点数 × 拉伸量`，K = 30.6（G2–G7 六组阶梯，误差 ≤ 0.1pt）
 *
 * 这个脚本重新从真值反推 2 与 3，并在「实现里的常数不再是最优解」时以退出码 1 失败。
 *
 *   node src/spike-compress.ts
 */
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { assertWindows } from './platform.ts';
import type { WordTruth } from './truth-types.ts';

assertWindows({ tool: '临时挤压穿刺', needs: '样本要用 Word COM 生成（真值本身跨平台可读）' });

const APP_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

/** 与 `@uw/layout` 的 `PUNCT_COMPRESS_MAX_EM` / `PUNCT_COMPRESS_STRETCH_K` 是同一对数 */
const MAX_EM = 0.48;
const K = 30.6;

interface Spec {
  paragraphs: { text: string; align: string; rightIndentPt?: number; sizePt: number }[];
}

/** 一段在真值里的落点：首行收了几个字、这一段一共几个字 */
interface Landing {
  text: string;
  align: string;
  rightIndentPt: number;
  total: number;
  onFirstLine: number;
  punct: number;
}

async function load(name: string): Promise<{ spec: Spec; truth: WordTruth }> {
  const spec = JSON.parse(
    await readFile(path.join(APP_ROOT, 'fixtures', 'src', `${name}.json`), 'utf8'),
  ) as Spec;
  const truth = JSON.parse(
    await readFile(path.join(APP_ROOT, 'fixtures', `${name}.truth.json`), 'utf8'),
  ) as WordTruth;
  return { spec, truth };
}

/**
 * 把真值的行按顺序摊回段落，只取**首行收了几个字**。
 *
 * 判据故意不依赖「这一段本来打算排几个字」：首行收了 m 个字，就等于 Word
 * **接受**了「m 个字要挤 x」，同时**拒绝**了「m+1 个字要挤 y」——
 * 一段给两个数据点，G10 那种后面还跟着两行的段落也照样能用。
 */
function landings(spec: Spec, truth: WordTruth): Landing[] {
  const lines = truth.pages.flatMap((p) => p.lines);
  const out: Landing[] = [];
  let i = 0;
  for (const p of spec.paragraphs) {
    let acc = '';
    let onFirstLine = 0;
    while (acc.length < p.text.length) {
      const line = lines[i++];
      if (line === undefined) throw new Error(`段落「${p.text.slice(0, 8)}…」在真值里断了`);
      if (acc === '') onFirstLine = [...line.text].length;
      acc += line.text;
    }
    if (acc !== p.text) throw new Error(`段落拼不回原文：${JSON.stringify(acc.slice(0, 20))}`);
    out.push({
      text: p.text,
      align: p.align,
      rightIndentPt: p.rightIndentPt ?? 0,
      total: [...p.text].length,
      onFirstLine,
      punct: [...p.text].filter((c) => c === '，').length,
    });
  }
  return out;
}

/** 逐字推进宽：从第一段（不缩进、无标点、绝不需要挤压）量，别拿 16.00 的标称值当基准 */
function perCharOf(truth: WordTruth): number {
  const first = truth.pages[0]?.lines[0];
  if (first === undefined) throw new Error('真值里没有行');
  return (first.xEnd - first.x) / [...first.text].length;
}

function contentWidth(truth: WordTruth): number {
  const s = truth.sections?.[0];
  if (s === undefined) throw new Error('真值里没有节的页面设置');
  return s.pageWidth - s.leftMargin - s.rightMargin;
}

let failed = false;

// ── ① 左对齐一格都不挤 ────────────────────────────────────────────────────────
{
  const { spec, truth } = await load('spike-compress-01');
  const rows = landings(spec, truth).filter((r) => r.align === 'left' && r.rightIndentPt > 0);
  const kept = rows.filter((r) => r.onFirstLine === r.total);
  console.log(`左对齐 ${rows.length} 段：留住最后一个字的有 ${kept.length} 段`);
  if (kept.length > 0) {
    console.error('✗ 左对齐居然挤了 —— `LineBreakContext.justified` 那条规则要重测');
    failed = true;
  } else {
    console.log('✓ 一段都没挤：临时挤压是两端对齐才有的行为\n');
  }
}

// ── ② / ③ 上限与兑换率 ───────────────────────────────────────────────────────
{
  const { spec, truth } = await load('spike-compress-02');
  const perChar = perCharOf(truth);
  const content = contentWidth(truth);
  const rows = landings(spec, truth).filter((r) => r.rightIndentPt > 0);

  /**
   * 模型：留住第 n 个字要挤 `need`，两条约束都要过。
   *
   * ① 容量：一个标点最多让出 `MAX_EM`；
   * ② 划算：`挤压量 × 字距数 ≤ K × 标点数 × 换行后要拉开的量`。
   */
  const wouldKeep = (need: number, punct: number, chars: number, avail: number): boolean => {
    if (need <= 0) return true;
    if (punct === 0) return false;
    if (need > punct * MAX_EM * perChar) return false;
    const stretch = avail - (chars - 1) * perChar;
    return need * (chars - 1) <= K * punct * stretch;
  };

  console.log(`逐字推进宽 ${perChar.toFixed(3)}pt，版心 ${content.toFixed(2)}pt`);
  console.log('每段给两个数据点：Word 接受了「收 m 个字」，同时拒绝了「收 m+1 个字」\n');
  console.log(
    `${'标点'.padStart(4)}${'首行字数'.padStart(10)}${'接受挤'.padStart(9)}${'拒绝挤'.padStart(9)}${'预测'.padStart(8)}`,
  );
  console.log('-'.repeat(42));

  let checked = 0;
  let wrong = 0;
  for (const r of rows) {
    const avail = content - r.rightIndentPt;
    const m = r.onFirstLine;
    const punctIn = (n: number): number => [...r.text].slice(0, n).filter((c) => c === '，').length;
    const accepted = m * perChar - avail;
    const refused = (m + 1) * perChar - avail;
    const okAccept = wouldKeep(accepted, punctIn(m), m, avail);
    const okRefuse = !wouldKeep(refused, punctIn(m + 1), m + 1, avail);
    checked += 2;
    if (!okAccept || !okRefuse) wrong++;
    if (!okAccept || !okRefuse) {
      console.log(
        String(r.punct).padStart(4) +
          String(m).padStart(10) +
          accepted.toFixed(2).padStart(9) +
          refused.toFixed(2).padStart(9) +
          `${okAccept ? '' : ' 该挤没挤'}${okRefuse ? '' : ' 该断没断'}`,
      );
    }
  }
  if (wrong === 0) {
    console.log(`${rows.length} 段 / ${checked} 个判断，模型全部说得通`);
  } else {
    console.error(`✗ ${wrong} 段与模型不符（上面列出）`);
    failed = true;
  }

  // 每组阶梯的翻转点：把 K 与上限的取值范围直接打出来，改常数时看这张表
  const groups = new Map<string, { need: number; kept: boolean }[]>();
  for (const r of rows) {
    const avail = content - r.rightIndentPt;
    const chars = [...r.text].length;
    // 只有「整段恰好一行」的阶梯才好按组读，G10 那种带尾巴的略过
    if (chars > r.onFirstLine + 1) continue;
    const key = `${r.punct}|${chars}`;
    const g = groups.get(key) ?? [];
    g.push({ need: chars * perChar - avail, kept: r.onFirstLine === chars });
    groups.set(key, g);
  }
  console.log(
    `\n${'标点'.padStart(4)}${'字数'.padStart(6)}${'实测翻转点'.padStart(16)}${'预测'.padStart(10)}${'谁卡住'.padStart(9)}`,
  );
  console.log('-'.repeat(48));
  for (const [key, g] of [...groups].sort()) {
    const [punct = 0, chars = 0] = key.split('|').map(Number);
    const lastKept = Math.max(...g.filter((x) => x.kept).map((x) => x.need));
    const firstDropped = Math.min(...g.filter((x) => !x.kept).map((x) => x.need));
    const byCap = punct * MAX_EM * perChar;
    const byRatio = (K * punct * perChar) / (chars - 1 + K * punct);
    const predicted = Math.min(byCap, byRatio);
    console.log(
      String(punct).padStart(4) +
        String(chars).padStart(6) +
        `${lastKept.toFixed(2)} → ${firstDropped.toFixed(2)}`.padStart(16) +
        predicted.toFixed(2).padStart(10) +
        (byCap < byRatio ? '上限' : '兑换率').padStart(8),
    );
  }
}
console.log(
  failed
    ? '\n✗ 未通过：实现里的 PUNCT_COMPRESS_MAX_EM / PUNCT_COMPRESS_STRETCH_K 与真值对不上'
    : `\n✓ 通过：一个标点最多让 ${MAX_EM} em，挤压与拉伸的兑换率 K = ${K}` +
        '（实现在 @uw/layout 的 roomOf / worthCompressing）',
);
if (failed) process.exit(1);
