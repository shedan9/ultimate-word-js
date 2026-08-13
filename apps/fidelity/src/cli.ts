#!/usr/bin/env node
/**
 * 真值流水线入口：
 *
 *   fixtures/src/<name>.json  --(Word COM)-->  fixtures/<name>.docx
 *   fixtures/<name>.docx      --(Word COM)-->  out/<name>.pdf + out/<name>.wordmeta.json
 *   out/<name>.pdf            --(pdf.js)  -->  fixtures/<name>.truth.json   ← 提交进仓库
 *
 * CI 上没有 Word，所以 truth.json 入库、pdf 不入库；本地跑 `pnpm truth` 重新生成。
 *
 * 用法：
 *   pnpm truth                 只处理过期的 fixture
 *   pnpm truth gongwen-01      只处理指定 fixture（可多个）
 *   pnpm truth --force         全部重跑
 *   pnpm truth --keep-pdf      保留中间 PDF（默认也保留在 out/，此开关关闭清理）
 */
import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { extractTruth } from './extract-truth.ts';
import { assertWindows } from './platform.ts';
import { runScript } from './run-powershell.ts';
import type { WordMeta } from './truth-types.ts';

const APP_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const FIXTURES = path.join(APP_ROOT, 'fixtures');
const SPECS = path.join(FIXTURES, 'src');
const OUT = path.join(APP_ROOT, 'out');
const SCRIPTS = path.join(APP_ROOT, 'scripts');

assertWindows({ tool: '真值流水线', needs: '导出 PDF 要调 Word 的 COM 接口' });

const args = process.argv.slice(2);
const force = args.includes('--force');
const only = new Set(args.filter((a) => !a.startsWith('--')));

const mtime = async (p: string): Promise<number> => {
  try {
    return (await stat(p)).mtimeMs;
  } catch {
    return Number.NaN; // 不存在
  }
};
const exists = async (p: string): Promise<boolean> => !Number.isNaN(await mtime(p));

const writeJson = async (p: string, value: unknown): Promise<void> => {
  // 确定性输出：无时间戳、固定缩进、末尾换行，diff 才有意义
  await writeFile(p, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
};

interface Fixture {
  name: string;
  spec?: string;
  docx: string;
  pdf: string;
  meta: string;
  truth: string;
}

async function discover(): Promise<Fixture[]> {
  const names = new Set<string>();
  for (const f of await readdir(FIXTURES).catch(() => [])) {
    if (f.toLowerCase().endsWith('.docx') && !f.startsWith('~$'))
      names.add(path.basename(f, path.extname(f)));
  }
  for (const f of await readdir(SPECS).catch(() => [])) {
    if (f.toLowerCase().endsWith('.json')) names.add(path.basename(f, '.json'));
  }
  const out: Fixture[] = [];
  for (const name of [...names].sort()) {
    if (only.size > 0 && !only.has(name)) continue;
    const spec = path.join(SPECS, `${name}.json`);
    out.push({
      name,
      ...((await exists(spec)) ? { spec } : {}),
      docx: path.join(FIXTURES, `${name}.docx`),
      pdf: path.join(OUT, `${name}.pdf`),
      meta: path.join(OUT, `${name}.wordmeta.json`),
      truth: path.join(FIXTURES, `${name}.truth.json`),
    });
  }
  return out;
}

async function build(fx: Fixture): Promise<boolean> {
  if (!fx.spec) return false;
  const [tSpec, tDocx] = [await mtime(fx.spec), await mtime(fx.docx)];
  if (!force && !Number.isNaN(tDocx) && tDocx >= tSpec) return false;
  await runScript(path.join(SCRIPTS, 'make-fixture.ps1'), ['-Spec', fx.spec, '-Output', fx.docx]);
  return true;
}

async function truth(fx: Fixture, rebuilt: boolean): Promise<'skipped' | 'ok'> {
  const [tDocx, tTruth] = [await mtime(fx.docx), await mtime(fx.truth)];
  if (Number.isNaN(tDocx)) throw new Error(`缺少 docx：${fx.docx}`);
  if (!force && !rebuilt && !Number.isNaN(tTruth) && tTruth >= tDocx) return 'skipped';

  await mkdir(OUT, { recursive: true });
  await runScript(path.join(SCRIPTS, 'export-truth.ps1'), [
    '-InputPath',
    fx.docx,
    '-OutputPdf',
    fx.pdf,
    '-MetaJson',
    fx.meta,
  ]);

  const meta = JSON.parse(await readFile(fx.meta, 'utf8')) as WordMeta;
  const result = await extractTruth(fx.pdf, { meta, source: path.basename(fx.docx) });
  await writeJson(fx.truth, result);

  if (result.wordPageCount !== undefined && result.wordPageCount !== result.pageCount) {
    console.warn(
      `  ! ${fx.name}: Word 报告 ${result.wordPageCount} 页，PDF 里是 ${result.pageCount} 页 —— 导出链路可疑`,
    );
  }
  const items = result.pages.reduce((n, p) => n + p.items.length, 0);
  const lines = result.pages.reduce((n, p) => n + p.lines.length, 0);
  console.log(
    `  ${fx.name}: ${result.pageCount} 页 / ${lines} 行 / ${items} 片段 → ${path.relative(APP_ROOT, fx.truth)}`,
  );
  return 'ok';
}

const fixtures = await discover();
if (fixtures.length === 0) {
  console.error(`没有找到 fixture（看过 ${path.relative(process.cwd(), FIXTURES)} 与它的 src/）`);
  process.exit(1);
}

let ok = 0;
let skipped = 0;
let failed = 0;
for (const fx of fixtures) {
  try {
    const rebuilt = await build(fx);
    if (rebuilt) console.log(`  ${fx.name}: 已从 spec 重建 docx`);
    if ((await truth(fx, rebuilt)) === 'skipped') {
      skipped++;
      console.log(`  ${fx.name}: 真值已是最新，跳过`);
    } else {
      ok++;
    }
  } catch (err) {
    failed++;
    console.error(`  ${fx.name}: 失败 —— ${err instanceof Error ? err.message : String(err)}`);
  }
}
console.log(`\n生成 ${ok} / 跳过 ${skipped} / 失败 ${failed}`);
process.exit(failed > 0 ? 1 : 0);
