/**
 * 生成标定样本用的纯色 PNG（`fixtures/src/square.png`）。
 *
 * 图片是二进制，扔进仓库就没人说得清它是怎么来的、里面到底有几个像素 ——
 * 而「图片本身多大」正是图片标定里最容易搞混的一件事（`wp:extent` 是**显示**尺寸，
 * 与像素数无关，样本要能证明这一点，就得知道像素数）。所以它由这个脚本产出：
 *
 *   node scripts/make-png.ts
 *
 * 8×8 纯黑、无 alpha —— 小到 PDF 里不可能被切片或分块，一张图就是一次绘制。
 */
import { writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { deflateSync } from 'node:zlib';

const SIZE = 8;

const crcTable = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(buf: Uint8Array): number {
  let c = 0xffffffff;
  for (const b of buf) c = ((crcTable[(c ^ b) & 0xff] as number) ^ (c >>> 8)) >>> 0;
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const body = Buffer.concat([Buffer.from(type, 'latin1'), data]);
  const out = Buffer.alloc(body.length + 8);
  out.writeUInt32BE(data.length, 0);
  body.copy(out, 4);
  out.writeUInt32BE(crc32(body), body.length + 4);
  return out;
}

const ihdr = Buffer.alloc(13);
ihdr.writeUInt32BE(SIZE, 0);
ihdr.writeUInt32BE(SIZE, 4);
ihdr[8] = 8; // 位深
ihdr[9] = 2; // 颜色类型 2 = truecolor RGB
// 每行前面那个 0 是 PNG 的「过滤器类型」字节，不是像素
const raw = Buffer.alloc(SIZE * (1 + SIZE * 3), 0);
for (let y = 0; y < SIZE; y++) raw[y * (1 + SIZE * 3)] = 0;

const png = Buffer.concat([
  Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
  chunk('IHDR', ihdr),
  chunk('IDAT', deflateSync(raw, { level: 9 })),
  chunk('IEND', new Uint8Array(0)),
]);

const out = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'fixtures', 'src', 'square.png');
writeFileSync(out, png);
console.log(`${out}  ${png.length} bytes  ${SIZE}x${SIZE}`);
