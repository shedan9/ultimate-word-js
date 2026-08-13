#!/usr/bin/env node
/**
 * Phase 0 穿刺：自己读字体表算出的单倍行距行高，能不能对上 Word 的实际基线？
 *
 * DoD（DEVELOPMENT-PLAN.md §4 Phase 0）：单页行基线误差 < 1pt。做不到就不要往下走。
 *
 * 做法：拿不开行网格的 fixture（网格会把基线吸到网格线上，掩盖字体度量的差异），
 * 按「同字体 + 同字号的连续行」分组，量相邻基线差，与 @uw/fonts 的预测对比。
 *
 *   node src/spike-lineheight.ts            # 跑默认的两组 spike fixture
 *   node src/spike-lineheight.ts <name>...  # 指定 fixture
 */
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { ptToTwips, twipsToPt } from '@uw/core';
import { lineMetrics, readRawMetrics } from '@uw/fonts';
import { hasEastAsianCoverage, openFont, WINDOWS_FONT_DIR } from '@uw/fonts/node';
import type { TruthPage, WordTruth } from './truth-types.ts';

const APP_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const DEFAULT_FIXTURES = ['spike-lineheight-01', 'spike-lineheight-02'];
/** DoD 阈值（pt） */
const TOLERANCE_PT = 1;

/**
 * PDF 里的字体名 → Windows 字体文件。
 * 这张表只服务于穿刺实验；正式的字体解析走 @uw/fonts 的替换表（Phase 2）。
 */
const FONT_FILES: Record<string, { file: string; postscriptName?: string }> = {
  FangSong: { file: 'simfang.ttf' },
  SimSun: { file: 'simsun.ttc', postscriptName: 'SimSun' },
  SimHei: { file: 'simhei.ttf' },
  KaiTi: { file: 'simkai.ttf' },
  MicrosoftYaHei: { file: 'msyh.ttc', postscriptName: 'MicrosoftYaHei' },
  DengXian: { file: 'Deng.ttf' },
  TimesNewRomanPSMT: { file: 'times.ttf' },
  ArialMT: { file: 'arial.ttf' },
};

interface FontEntry {
  metrics: ReturnType<typeof readRawMetrics>;
  eastAsian: boolean;
}

const fontCache = new Map<string, FontEntry>();

function loadFont(pdfName: string): FontEntry {
  const hit = fontCache.get(pdfName);
  if (hit) return hit;
  const entry = FONT_FILES[pdfName];
  if (!entry) throw new Error(`没有登记字体文件：${pdfName}（补进 spike-lineheight.ts 的 FONT_FILES）`);
  const font = openFont(
    path.join(WINDOWS_FONT_DIR, entry.file),
    ...(entry.postscriptName ? [entry.postscriptName] : []),
  );
  const loaded: FontEntry = { metrics: readRawMetrics(font), eastAsian: hasEastAsianCoverage(font) };
  fontCache.set(pdfName, loaded);
  return loaded;
}

/** 一行的「字体 + 字号」签名；同签名的连续行才能拿来量基线差 */
function lineSignature(page: TruthPage, lineIndex: number): string {
  const line = page.lines[lineIndex];
  if (!line) return '';
  const parts = new Set<string>();
  for (const i of line.items) {
    const it = page.items[i];
    if (it) parts.add(`${it.font}@${it.size}`);
  }
  return [...parts].sort().join(' + ');
}

/** 该行是否含东亚文字 —— 用字符判定，不看字体 */
const EAST_ASIAN_RE = /[⺀-鿿豈-﫿＀-｠]/;

/** 混排行的行高 = 各字体行高的最大值（ascent / descent 逐项取最大，这里只关心总高） */
function predictLineHeightPt(page: TruthPage, lineIndex: number): number {
  const line = page.lines[lineIndex];
  if (!line) return 0;
  const seen = new Set<string>();
  let maxTwips = 0;
  for (const i of line.items) {
    const it = page.items[i];
    if (!it) continue;
    const key = `${it.font}@${it.size}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const { metrics, eastAsian } = loadFont(it.font);
    // 字号取 PDF 里的实际值：Word 导出时会把 16pt 写成 15.96，用标称值反而引入误差
    const lm = lineMetrics(metrics, ptToTwips(it.size), {
      eastAsian: eastAsian && EAST_ASIAN_RE.test(line.text),
    });
    if (lm.lineHeight > maxTwips) maxTwips = lm.lineHeight;
  }
  return twipsToPt(maxTwips);
}

interface Sample {
  fixture: string;
  signature: string;
  measured: number;
  predicted: number;
  text: string;
}

async function collect(fixture: string): Promise<Sample[]> {
  const truth = JSON.parse(
    await readFile(path.join(APP_ROOT, 'fixtures', `${fixture}.truth.json`), 'utf8'),
  ) as WordTruth;

  const samples: Sample[] = [];
  for (const page of truth.pages) {
    for (let i = 1; i < page.lines.length; i++) {
      const prev = page.lines[i - 1];
      const cur = page.lines[i];
      if (!prev || !cur) continue;
      // 只在同签名的相邻行之间量：跨段落的基线差里混着段前段后间距，不是行高
      const sig = lineSignature(page, i);
      if (sig !== lineSignature(page, i - 1)) continue;
      samples.push({
        fixture,
        signature: sig,
        measured: cur.y - prev.y,
        predicted: predictLineHeightPt(page, i),
        text: cur.text.slice(0, 16),
      });
    }
  }
  return samples;
}

const names = process.argv.slice(2).filter((a) => !a.startsWith('--'));
const fixtures = names.length > 0 ? names : DEFAULT_FIXTURES;

const samples: Sample[] = [];
for (const f of fixtures) samples.push(...(await collect(f)));

if (samples.length === 0) {
  console.error('没有可比对的样本 —— fixture 里没有同字体同字号的相邻行？');
  process.exit(1);
}

console.log('Phase 0 穿刺 · 单倍行距行高（不开行网格）\n');
console.log(`${'字体 @ 字号'.padEnd(40)}${'实测'.padStart(9)}${'预测'.padStart(9)}${'误差'.padStart(9)}`);
console.log('-'.repeat(67));

let worst = 0;
let worstSample: Sample | undefined;
for (const s of samples) {
  const err = s.measured - s.predicted;
  if (Math.abs(err) > Math.abs(worst)) {
    worst = err;
    worstSample = s;
  }
  console.log(
    `${s.signature.padEnd(40)}${s.measured.toFixed(2).padStart(9)}${s.predicted.toFixed(2).padStart(9)}${err.toFixed(3).padStart(9)}`,
  );
}

console.log('-'.repeat(67));
console.log(`样本 ${samples.length} 行，最大误差 ${worst.toFixed(3)} pt（阈值 ${TOLERANCE_PT} pt）`);
if (worstSample)
  console.log(
    `最差样本：${worstSample.fixture} · ${worstSample.signature} · ${JSON.stringify(worstSample.text)}`,
  );

if (Math.abs(worst) > TOLERANCE_PT) {
  console.error('\n✗ Phase 0 未通过：行高算法对不上 Word，先解决它再往下走');
  process.exit(1);
}
console.log('\n✓ Phase 0 通过：读 OS/2 表算行高，单页基线误差在阈值内');
