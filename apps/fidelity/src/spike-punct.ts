#!/usr/bin/env node
/**
 * 标点挤压穿刺：Word 到底什么时候压标点、压多少。
 *
 * 起因是 gongwen-01 的真值对不上：我们只在「行尾塞不下」时挤，Word 的行却处处比我们窄。
 * 但那份正文一行里同时叠着挤压、悬挂、中西文间距三件事，反推不出干净的系数 ——
 * 所以另做一份 `spike-punct-01`：**每段只放一个孤立的标点或一对相邻标点，且短到绝不折行**，
 * 于是「行宽 − 自然宽」就只剩挤压这一个变量。
 *
 * 结论（把最初的猜测推翻了）：**孤立的标点一点都不压，只有标点紧跟标点才压，固定半个字**。
 * 断言就照这条写：每一行的压缩量必须等于 `0.5 em × 相邻标点对数`。
 *
 *   node src/spike-punct.ts
 */
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { assertWindows } from './platform.ts';
import type { WordTruth } from './truth-types.ts';

assertWindows({ tool: '标点挤压穿刺', needs: '样本要用 Word COM 生成（真值本身跨平台可读）' });

const APP_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const FIXTURE = 'spike-punct-01';
/** 样本里所有段落都是这个字号（spec 写死的），标称值，不取 PDF 的量化值 */
const SIZE_PT = 16;
/** 每对相邻标点压掉多少 em —— 与 `@uw/layout` 的 `PUNCT_PAIR_COMPRESS_EM` 是同一个数 */
const PAIR_EM = 0.5;
/** 容差（em）。实测残差 0.005 em 量级，来自 PDF 坐标量化 */
const TOLERANCE_EM = 0.02;

/**
 * 可挤压的全角标点 —— 与 `@uw/layout` 的 `break-class.ts` 那张表**必须一致**。
 * 这里刻意重抄一遍而不是 import：fidelity 不依赖 layout（依赖方向单向），
 * 而两张表不一致的后果正是这个脚本要抓的东西。
 */
const COMPRESSIBLE = new Set([...'、。，．：；！？「」『』（）〔〕【】《》〈〉“”‘’']);

/** 一行里有多少对「标点紧跟标点」 */
function pairsIn(text: string): number {
  const chars = [...text];
  let pairs = 0;
  for (let i = 1; i < chars.length; i++) {
    if (COMPRESSIBLE.has(chars[i - 1] as string) && COMPRESSIBLE.has(chars[i] as string)) pairs++;
  }
  return pairs;
}

const truth = JSON.parse(
  await readFile(path.join(APP_ROOT, 'fixtures', `${FIXTURE}.truth.json`), 'utf8'),
) as WordTruth;
const page = truth.pages[0];
if (!page) throw new Error(`${FIXTURE}: 真值里没有页`);

/**
 * 每个汉字的实际推进宽度，从**没有标点的那一段**量出来。
 *
 * 不直接用 16.00pt：Word 导出 PDF 时把字号写成 15.96，逐字推进宽实测 15.982 ——
 * 拿标称值当基准，5 个字就凑出 0.09pt 的假压缩量，与真实信号（8pt）比虽小，
 * 但会让「孤立标点压 0.005 em」这种噪声看起来像结论。
 */
const control = page.lines.find((l) => [...l.text].every((ch) => !COMPRESSIBLE.has(ch)));
if (!control) throw new Error(`${FIXTURE}: 找不到不含标点的对照段落`);
const perChar = (control.xEnd - control.x) / [...control.text].length;

console.log(`标点挤压穿刺 · ${FIXTURE}（${SIZE_PT}pt 仿宋，逐字推进宽实测 ${perChar.toFixed(3)}pt）\n`);
console.log(
  `${'段落文字'.padEnd(20)}${'字数'.padStart(5)}${'压掉 em'.padStart(9)}${'相邻对数'.padStart(9)}${'预期 em'.padStart(9)}${'误差'.padStart(9)}`,
);
console.log('-'.repeat(62));

let worst = 0;
let worstText = '';
for (const line of page.lines) {
  const chars = [...line.text];
  const measured = (chars.length * perChar - (line.xEnd - line.x)) / SIZE_PT;
  const pairs = pairsIn(line.text);
  const expected = pairs * PAIR_EM;
  const err = measured - expected;
  if (Math.abs(err) > Math.abs(worst)) {
    worst = err;
    worstText = line.text;
  }
  console.log(
    line.text.padEnd(20) +
      String(chars.length).padStart(5) +
      measured.toFixed(4).padStart(9) +
      String(pairs).padStart(9) +
      expected.toFixed(4).padStart(9) +
      err.toFixed(4).padStart(9),
  );
}

console.log('-'.repeat(62));
console.log(`${page.lines.length} 段，最大误差 ${worst.toFixed(4)} em（阈值 ${TOLERANCE_EM}）`);
if (worstText !== '') console.log(`最差样本：${worstText}`);

if (Math.abs(worst) > TOLERANCE_EM) {
  console.error('\n✗ 未通过：压缩量不等于「0.5 em × 相邻标点对数」，规则要重测');
  process.exit(1);
}
console.log('\n✓ 通过：孤立标点不压，相邻标点固定压半个字（实现在 @uw/layout 的 applyPunctPairs）');
