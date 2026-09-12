/**
 * 随库那 17 款度量包的**无 fs** 入口 —— 浏览器与 Node 同一条路。
 *
 * `@uw/fonts/node` 的 `loadBundledPacks()` 靠 `readFileSync` 读 `packs/`，浏览器走不了；
 * 而门面包承诺「`load()` 时替调用方把随库字体做掉」（api.md §4），所以得有一条
 * 打包器能跟着走的路。JSON import attributes（`with { type: 'json' }`）是唯一同时被
 * Node 24、Vite 与 TypeScript 认的写法；`import.meta.glob` 是 Vite 专属，不能进库。
 *
 * 这批 JSON 是**数据不是模块**，所以一律先收成 `unknown` 再过 `asPack()`：
 * 让 TS 直接推 JSON 的字面量类型会把 `version: 1` 推成 `number`、
 * 又把 88 KB 的数字表焊进类型检查，两头都不划算。
 *
 * 列表是**手写的**（`tools/build-packs.ts` 只在 Windows 上跑，生成不了这个文件），
 * 由 `packs.test.ts` 对着 `packs/index.json` 校验 —— 抽了新包忘了加进来会当场红。
 */

import arial from '../packs/Arial.json' with { type: 'json' };
import calibri from '../packs/Calibri.json' with { type: 'json' };
import cambria from '../packs/Cambria.json' with { type: 'json' };
import courierNew from '../packs/Courier New.json' with { type: 'json' };
import georgia from '../packs/Georgia.json' with { type: 'json' };
import segoeUi from '../packs/Segoe UI.json' with { type: 'json' };
import symbol from '../packs/Symbol.json' with { type: 'json' };
import tahoma from '../packs/Tahoma.json' with { type: 'json' };
import timesNewRoman from '../packs/Times New Roman.json' with { type: 'json' };
import verdana from '../packs/Verdana.json' with { type: 'json' };
import wingdings from '../packs/Wingdings.json' with { type: 'json' };
import fangsong from '../packs/仿宋.json' with { type: 'json' };
import simsun from '../packs/宋体.json' with { type: 'json' };
import yahei from '../packs/微软雅黑.json' with { type: 'json' };
import kaiti from '../packs/楷体.json' with { type: 'json' };
import dengxian from '../packs/等线.json' with { type: 'json' };
import simhei from '../packs/黑体.json' with { type: 'json' };
import type { MetricsPack } from './metrics-pack.ts';

function asPack(data: unknown): MetricsPack {
  const p = data as Partial<MetricsPack>;
  if (p.version !== 1 || typeof p.family !== 'string') {
    throw new Error(`随库度量包格式不对：${JSON.stringify(data).slice(0, 80)}`);
  }
  return p as MetricsPack;
}

const RAW: readonly unknown[] = [
  simsun,
  fangsong,
  simhei,
  kaiti,
  dengxian,
  yahei,
  timesNewRoman,
  arial,
  calibri,
  cambria,
  courierNew,
  georgia,
  segoeUi,
  symbol,
  tahoma,
  verdana,
  wingdings,
];

/** 随库度量包。每次调用返回同一批对象 —— 度量包是只读数据，没有理由复制 */
export function bundledPacks(): readonly MetricsPack[] {
  return PACKS;
}

const PACKS: readonly MetricsPack[] = RAW.map(asPack);
