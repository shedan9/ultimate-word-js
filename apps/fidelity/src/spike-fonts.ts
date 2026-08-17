/**
 * 穿刺实验专用的「PDF 字体名 → Windows 字体文件」表。
 *
 * 只服务于保真度实验：真值里的字体名是 pdf.js 从 PDF 字体表读出来的 PostScript 名
 * （`TimesNewRomanPSMT`），而正式的字体解析走 `@uw/model` 的 `fontNameCandidates()`
 * ——文档里写的是「仿宋」这种本地化名，两条路的入口不同，不要混用。
 *
 * 表是**显式**的而不是去扫字体目录：扫出来的东西会随机器变，
 * 而穿刺实验的结论要能在另一台机器上复算。缺登记就报错，不猜。
 */
import path from 'node:path';
import type { RawFontMetrics } from '@uw/fonts';
import { readRawMetrics } from '@uw/fonts';
import { hasEastAsianCoverage, openFont, WINDOWS_FONT_DIR } from '@uw/fonts/node';

interface FontFile {
  file: string;
  /** `.ttc` 字体集里要指定哪一款（simsun.ttc 里同时有 SimSun 与 NSimSun） */
  postscriptName?: string;
}

const FONT_FILES: Record<string, FontFile> = {
  FangSong: { file: 'simfang.ttf' },
  SimSun: { file: 'simsun.ttc', postscriptName: 'SimSun' },
  SimHei: { file: 'simhei.ttf' },
  KaiTi: { file: 'simkai.ttf' },
  MicrosoftYaHei: { file: 'msyh.ttc', postscriptName: 'MicrosoftYaHei' },
  DengXian: { file: 'Deng.ttf' },
  TimesNewRomanPSMT: { file: 'times.ttf' },
  ArialMT: { file: 'arial.ttf' },
};

export interface SpikeFont {
  metrics: RawFontMetrics;
  /** 该字体是否覆盖东亚文字 —— 决定它算不算「定行盒的那一款」 */
  eastAsian: boolean;
}

const cache = new Map<string, SpikeFont>();

export function loadSpikeFont(pdfName: string): SpikeFont {
  const hit = cache.get(pdfName);
  if (hit) return hit;
  const entry = FONT_FILES[pdfName];
  if (!entry) throw new Error(`没有登记字体文件：${pdfName}（补进 spike-fonts.ts 的 FONT_FILES）`);
  const font = openFont(
    path.join(WINDOWS_FONT_DIR, entry.file),
    ...(entry.postscriptName ? [entry.postscriptName] : []),
  );
  const loaded: SpikeFont = { metrics: readRawMetrics(font), eastAsian: hasEastAsianCoverage(font) };
  cache.set(pdfName, loaded);
  return loaded;
}

/**
 * 一行里有没有东亚文字 —— **按字符判定，不看字体**。
 *
 * 这与 `@uw/layout` 的 `hasEastAsia()` 同构：Word 是按行选行距规则的，
 * 一行里有一个汉字，整行就走东亚那套。
 */
const EAST_ASIAN_RE = /[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffe0]/;

export function hasEastAsianText(text: string): boolean {
  return EAST_ASIAN_RE.test(text);
}
