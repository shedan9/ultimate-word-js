// 用 headless Chrome + CDP 跑这里的浏览器回归页，轮询每页的 #result 直到 PASS / FAIL。
//
// 为什么不直接在浏览器里开：回归页靠 rAF 等一帧，被遮住的标签页（document.hidden）里 rAF 不跑，
// 页面会永远停在「正在验证…」；headless 没有这个问题，而且能一口气跑完七个页面。
// 用法（先 `pnpm --filter @uw/playground dev`）：
//   node apps/playground/tests/run.mjs            # 全部
//   node apps/playground/tests/run.mjs print react  # 指定页
// Chrome 路径可用 CHROME 环境变量覆盖；靠 Node 24 原生 WebSocket，不装任何依赖。
import { spawn } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const CHROME = process.env.CHROME ?? '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const BASE = process.env.PLAYGROUND ?? 'http://localhost:5273';
const ALL = ['view', 'virtual-text', 'annotations', 'scroll', 'facade', 'print', 'react', 'editing'];
const pages = process.argv.length > 2 ? process.argv.slice(2) : ALL;
const port = 9333;
const profile = mkdtempSync(join(tmpdir(), 'uw-chrome-'));
const chrome = spawn(
  CHROME,
  [
    '--headless=new',
    '--disable-gpu',
    '--no-first-run',
    `--remote-debugging-port=${port}`,
    `--user-data-dir=${profile}`,
    '--window-size=1400,900',
    'about:blank',
  ],
  { stdio: 'ignore' },
);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function ready() {
  for (let i = 0; i < 50; i++) {
    try {
      if ((await fetch(`http://127.0.0.1:${port}/json/version`)).ok) return;
    } catch {}
    await sleep(200);
  }
  throw new Error(`Chrome 没起来（${CHROME}）`);
}

async function run(url) {
  const target = await (await fetch(`http://127.0.0.1:${port}/json/new?${url}`, { method: 'PUT' })).json();
  const ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((r) => {
    ws.onopen = r;
  });
  let id = 0;
  const pending = new Map();
  ws.onmessage = (e) => {
    const m = JSON.parse(e.data);
    if (m.id && pending.has(m.id)) {
      pending.get(m.id)(m.result);
      pending.delete(m.id);
    }
  };
  const send = (method, params = {}) =>
    new Promise((r) => {
      const n = ++id;
      pending.set(n, r);
      ws.send(JSON.stringify({ id: n, method, params }));
    });
  const evaluate = async (expression) =>
    (await send('Runtime.evaluate', { expression, returnByValue: true })).result?.value;
  const t0 = Date.now();
  let text = '';
  while (Date.now() - t0 < 60000) {
    text = (await evaluate("document.querySelector('#result')?.textContent ?? ''")) ?? '';
    if (/^(PASS|FAIL)/.test(text)) break;
    await sleep(250);
  }
  ws.close();
  await fetch(`http://127.0.0.1:${port}/json/close/${target.id}`);
  return text;
}

try {
  await ready();
  let failed = false;
  for (const page of pages) {
    const text = await run(`${BASE}/tests/${page}.html`);
    console.log(`${page.padEnd(14)} ${text.startsWith('PASS') ? text.split('\n')[0] : text}`);
    if (!text.startsWith('PASS')) failed = true;
  }
  process.exitCode = failed ? 1 : 0;
} finally {
  // Chrome 退出时还在往 profile 里写，等它真退了再删
  const exited = new Promise((r) => chrome.once('exit', r));
  chrome.kill();
  await Promise.race([exited, sleep(3000)]);
  rmSync(profile, { recursive: true, force: true, maxRetries: 5 });
}
