#!/usr/bin/env node
/**
 * 度量包抽取 —— 从 `C:/Windows/Fonts` 抽出随库分发的纯度量 JSON。
 *
 *   pnpm --filter @uw/fonts run packs           # 抽全部，写进 packs/
 *   pnpm --filter @uw/fonts run packs 宋体 黑体  # 只抽指定的（按包名或字体名）
 *   node tools/build-packs.ts --check           # 只检查现有包是否与本机字体一致，不写
 *
 * **只能在 Windows 上跑**，而且必须是 `C:/Windows/Fonts` 里的那一份：包的意义是把
 * Word 用的那套度量搬到别的平台去。抽 macOS 上的同名字体等于把误差固化进随库分发的文件里
 * —— 那比没有包更糟，因为它看起来是对的。
 *
 * 产物入库（`packs/*.json`），所以**消费侧完全跨平台**：CI 与 Mac 上跑测试读的就是这些文件，
 * 这是坐标级真值断言（L2 断行点 / L4 片段 x）能在非 Windows 机器上跑起来的前提。
 * 与真值 `*.truth.json` 同一个套路：生成绑 Windows，消费不绑。
 */
import { mkdir, readdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildPackFromFile, WINDOWS_FONT_DIR } from '../src/load-node.ts';
import type { MetricsPack } from '../src/metrics-pack.ts';
import { symbolSampleCodePoints } from '../src/metrics-pack.ts';

const PACKS_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'packs');

interface FontEntry {
  /** 包文件名（不含扩展名），同时是 `MetricsPack.family` —— 用**文档里会出现的那个名字** */
  family: string;
  file: string;
  /** `.ttc` 字体集里指定哪一款（simsun.ttc 里同时有 SimSun 与 NSimSun） */
  postscriptName?: string;
  /** A / B / C / D 四类，见开发计划 §2.1 的字体清单 */
  klass: 'A' | 'B' | 'C' | 'D';
  /** symbol-encoded 字体要采满 0x00–0xFF，见 `symbolSampleCodePoints()` */
  symbol?: true;
}

/**
 * 字体清单。**显式列表，不扫目录** —— 扫出来的东西随机器变，
 * 而入库的产物必须能在另一台装了同版本 Word 的机器上复算出同一份。
 *
 * `family` 用的是**中文名**（A 类）：中文版 Word 在 `w:rFonts` 里写的就是「仿宋」，
 * 磁盘上的文件叫 `simfang.ttf`、字体自报 PostScript 名 `FangSong`。注册表按 `family` 建索引，
 * 而 `@uw/model` 的 `fontNameCandidates()` 会把 `w:altName` 也算进候选，
 * 所以两个名字都能命中 —— 但索引键取文档里那个才是主路径。
 */
const FONTS: FontEntry[] = [
  // A 类 · 中文：没有度量兼容的开源替代，只能抽度量，渲染时用 Noto / 文楷顶上
  { family: '宋体', file: 'simsun.ttc', postscriptName: 'SimSun', klass: 'A' },
  { family: '仿宋', file: 'simfang.ttf', klass: 'A' },
  { family: '黑体', file: 'simhei.ttf', klass: 'A' },
  { family: '楷体', file: 'simkai.ttf', klass: 'A' },
  // 等线与微软雅黑不在开发计划的首批 A 类表里，但公文之外的周报 / 报告绕不开它们：
  // 等线是 Office 2016+ 中文版的默认正文字体，微软雅黑是默认标题 / UI 字体。
  // 两者的 win 跨度（1.0420 / 1.3198 em）与宋体家族（1.0000 em）差得远，
  // 缺了包就会在行高上直接错三成，不是精度问题而是错得看得见。
  { family: '等线', file: 'Deng.ttf', klass: 'A' },
  { family: '微软雅黑', file: 'msyh.ttc', postscriptName: 'MicrosoftYaHei', klass: 'A' },
  // B 类 · 拉丁核心：最终方案是打包度量兼容的开源克隆（Liberation / Carlito / Caladea），
  // 那样字形也一并解决。克隆还没进仓库，先抽包顶上 —— 包只保证**排版**一致，
  // 字形仍会退到系统 serif / sans。克隆进来之后这几行可以删。
  { family: 'Times New Roman', file: 'times.ttf', klass: 'B' },
  { family: 'Arial', file: 'arial.ttf', klass: 'B' },
  { family: 'Courier New', file: 'cour.ttf', klass: 'B' },
  { family: 'Calibri', file: 'calibri.ttf', klass: 'B' },
  { family: 'Cambria', file: 'cambria.ttc', postscriptName: 'Cambria', klass: 'B' },
  // C 类 · 拉丁次要：只给度量，渲染回退系统 serif / sans
  { family: 'Georgia', file: 'georgia.ttf', klass: 'C' },
  { family: 'Verdana', file: 'verdana.ttf', klass: 'C' },
  { family: 'Tahoma', file: 'tahoma.ttf', klass: 'C' },
  { family: 'Segoe UI', file: 'segoeui.ttf', klass: 'C' },
  // D 类 · 符号字体：项目符号的实际载体（numbering.xml 里实心圆点是 Symbol 的 0xB7）。
  // 缺了它所有列表 bullet 的宽度都是猜的，而 bullet 宽度决定首行正文从哪儿开始
  { family: 'Symbol', file: 'symbol.ttf', klass: 'D', symbol: true },
  { family: 'Wingdings', file: 'wingding.ttf', klass: 'D', symbol: true },
];

if (process.platform !== 'win32') {
  console.error(
    [
      `度量包抽取只能在 Windows 上跑：要读 ${WINDOWS_FONT_DIR} 里 Word 实际使用的那一份字体。`,
      `当前平台是 ${process.platform}。`,
      '',
      '包的**消费**不需要 Windows —— packs/*.json 已经入库，直接读就行：',
      '  pnpm turbo run test',
      '',
      '只有新增字体或换了 Word / 系统版本时才需要回 Windows 重抽一次。',
    ].join('\n'),
  );
  // 退出码 2 = 环境不满足，区别于 1 = 抽取失败。与 apps/fidelity 的 platform.ts 一致
  process.exit(2);
}

const args = process.argv.slice(2);
const check = args.includes('--check');
const only = new Set(args.filter((a) => !a.startsWith('--')));

/**
 * 确定性输出：字段顺序固定、不写时间戳，否则每次重抽 diff 全是噪声。
 *
 * 排版是**每个顶层字段一行**，而不是 `JSON.stringify(x, null, 2)`：后者会把 coverage 的
 * 每个区间摊成三行，等线一款就从 8KB 涨到 16.6KB，17 个包白多出 100KB。
 * 逐行 diff 对这种文件也没意义 —— 它是生成数据，变化只可能是「字体文件换了」，
 * 一行一字段刚好够看出「变的是度量还是覆盖范围」。
 */
const NL = '\n';

function serialize(pack: MetricsPack): string {
  const fields = Object.entries(pack).map(([k, v]) => `  ${JSON.stringify(k)}: ${JSON.stringify(v)}`);
  return `{${NL}${fields.join(`,${NL}`)}${NL}}${NL}`;
}

await mkdir(PACKS_DIR, { recursive: true });

let written = 0;
let unchanged = 0;
let failed = 0;
const built = new Map<string, MetricsPack>();

for (const entry of FONTS) {
  if (only.size > 0 && !only.has(entry.family) && !only.has(entry.file)) continue;
  const source = path.join(WINDOWS_FONT_DIR, entry.file);
  try {
    const pack = buildPackFromFile(source, entry.postscriptName, {
      family: entry.family,
      ...(entry.symbol === true ? { sample: symbolSampleCodePoints() } : {}),
    });
    built.set(entry.family, pack);

    const target = path.join(PACKS_DIR, `${entry.family}.json`);
    const next = serialize(pack);
    const prev = await readFile(target, 'utf8').catch(() => undefined);
    if (prev === next) {
      unchanged++;
      continue;
    }
    if (check) {
      failed++;
      console.error(
        `  ✗ ${entry.family}：入库的包与本机字体${prev === undefined ? '缺失' : '不一致'} —— 跑一次 pnpm packs`,
      );
      continue;
    }
    await writeFile(target, next, 'utf8');
    written++;
    const coverage = pack.coverage?.length ?? 0;
    console.log(
      `  ${entry.klass} ${entry.family.padEnd(16)} em=${String(pack.unitsPerEm).padStart(4)}` +
        ` 默认宽=${String(pack.defaultAdvance).padStart(4)}` +
        ` 例外=${String(Object.keys(pack.advances).length).padStart(3)}` +
        ` 覆盖区间=${String(coverage).padStart(4)}` +
        ` ${(next.length / 1024).toFixed(1)}KB`,
    );
  } catch (err) {
    failed++;
    console.error(
      `  ✗ ${entry.family}（${entry.file}）：${err instanceof Error ? err.message : String(err)}`,
    );
  }
}

// 索引文件让消费侧不必把字体名硬编码一遍：浏览器 fetch、Node 读盘、打包器 glob 都用它
if (!check && built.size > 0) {
  const listed = FONTS.filter((f) => built.has(f.family)).map((f) => ({
    family: f.family,
    file: `${f.family}.json`,
    class: f.klass,
  }));
  const indexPath = path.join(PACKS_DIR, 'index.json');
  // 只抽了一部分时不要把索引截断成那一部分：读回旧索引合并
  const prevIndex = JSON.parse(await readFile(indexPath, 'utf8').catch(() => '[]')) as typeof listed;
  const merged = [...prevIndex.filter((p) => !built.has(p.family)), ...listed];
  const order = new Map(FONTS.map((f, i) => [f.family, i]));
  merged.sort((a, b) => (order.get(a.family) ?? 999) - (order.get(b.family) ?? 999));
  await writeFile(indexPath, `${JSON.stringify(merged, null, 2)}\n`, 'utf8');
}

const stale = (await readdir(PACKS_DIR).catch(() => []))
  .filter((f) => f.endsWith('.json') && f !== 'index.json')
  .map((f) => path.basename(f, '.json'))
  .filter((name) => !FONTS.some((f) => f.family === name));
if (stale.length > 0) {
  console.warn(`  ! packs/ 里有清单外的包：${stale.join('、')} —— 要么补进 FONTS，要么删掉`);
}

console.log(`\n写入 ${written} / 未变 ${unchanged} / 失败 ${failed}`);
process.exit(failed > 0 ? 1 : 0);
