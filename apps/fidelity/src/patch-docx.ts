/**
 * 造完 docx 之后**直接改 XML**：给相邻两格的共享边写上**互相冲突**的边框。
 *
 * ## 为什么非改 XML 不可
 *
 * Word 的对象模型里，一条共享边只有**一个** `Border` 对象：给左格设 `right`、
 * 再给右格设 `left`，第二次设的把第一次的整个盖掉，存出来的 `w:tcBorders` 两边
 * 一模一样（实测：12 组配对全部塌成同一条，见 `spike-table-03` 的 note）。
 * 也就是说 **Word 自己造不出「相邻竞争」这个局面** —— 竞争只发生在别的生成器写出来的、
 * 从别处粘进来的、或经 `w:tblPrEx` 改过的文件上。而我们要标定的恰恰是 Word **渲染**
 * 这类文件时的规则，所以样本只能这么造：Word 排版 + 手工写冲突 + Word 导 PDF。
 *
 * ## 定位靠文字不靠下标
 *
 * 补丁的 key 是**格子里的文字**（`一左` / `一右`），不是「第几张表第几行第几格」：
 * 下标要跟着 spec 的增删改动，而文字就在 spec 里摆着，对不上时报错也说得清是哪一格。
 * 同一份 fixture 里文字重复就报错 —— 悄悄打到第一个匹配上是最难查的那种错。
 */
import { readFile, writeFile } from 'node:fs/promises';
import { unzipSync, zipSync } from 'fflate';

/** 一条边的原始属性，与 `w:tcBorders` 的子元素一一对应 */
export interface RawBorder {
  /** `w:val`：`single` / `double` / `dashed` / `nil` … */
  val: string;
  /** `w:sz`，单位 1/8 磅。`nil` 时可省 */
  sz?: number;
  /** `w:color`，六位十六进制或 `auto` */
  color?: string;
}

export type RawCellBorders = Partial<
  Record<'top' | 'left' | 'bottom' | 'right' | 'insideH' | 'insideV', RawBorder>
>;

/** `w:tcBorders` 里子元素的顺序（CT_TcBorders 是 sequence，顺序错了 Word 会拒绝打开） */
const SIDE_ORDER = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const;

/**
 * `w:tcPr` 里排在 `w:tcBorders` **之后**的那些元素（CT_TcPr 也是 sequence）。
 * 插入时要插在它们之前 —— 直接塞在 `</w:tcPr>` 前面的话，带 `w:vAlign` 或 `w:tcMar`
 * 的格子就会写出顺序颠倒的 XML，Word 报「文件已损坏」而不是忽略它。
 */
const AFTER_BORDERS = [
  'w:shd',
  'w:noWrap',
  'w:tcMar',
  'w:textDirection',
  'w:tcFitText',
  'w:vAlign',
  'w:hideMark',
];

const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');

function bordersXml(borders: RawCellBorders): string {
  const parts: string[] = [];
  for (const side of SIDE_ORDER) {
    const b = borders[side];
    if (b === undefined) continue;
    const attrs = [`w:val="${esc(b.val)}"`];
    // nil 的那三个属性 Word 自己也不写；写上去不算错，但 diff 里多三样对不上的东西
    if (b.sz !== undefined) attrs.push(`w:sz="${b.sz}"`, 'w:space="0"');
    if (b.color !== undefined) attrs.push(`w:color="${esc(b.color)}"`);
    parts.push(`<w:${side} ${attrs.join(' ')}/>`);
  }
  return `<w:tcBorders>${parts.join('')}</w:tcBorders>`;
}

/** 扫出每个 `<w:tc>` 的区间（含嵌套表格里的），按开始位置排序 */
function cellRanges(xml: string): { start: number; end: number }[] {
  const re = /<w:tc(?:\s[^>]*)?>|<\/w:tc>/g;
  const open: number[] = [];
  const out: { start: number; end: number }[] = [];
  for (let m = re.exec(xml); m !== null; m = re.exec(xml)) {
    if (m[0].startsWith('</')) {
      const start = open.pop();
      if (start !== undefined) out.push({ start, end: m.index + m[0].length });
    } else {
      open.push(m.index);
    }
  }
  return out.sort((a, b) => a.start - b.start);
}

/** 格子里的文字：所有 `w:t` 拼起来，与 spec 里写的那串对得上 */
function cellText(chunk: string): string {
  let text = '';
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  for (let m = re.exec(chunk); m !== null; m = re.exec(chunk)) text += m[1] ?? '';
  return text;
}

/** 把一格的 `w:tcBorders` 换成给定的那份（没有就按 schema 顺序插进去） */
function applyToCell(chunk: string, borders: RawCellBorders): string {
  const xml = bordersXml(borders);
  if (/<w:tcBorders>[\s\S]*?<\/w:tcBorders>/.test(chunk)) {
    return chunk.replace(/<w:tcBorders>[\s\S]*?<\/w:tcBorders>/, xml);
  }
  const pr = /<w:tcPr>([\s\S]*?)<\/w:tcPr>/.exec(chunk);
  if (pr === null) throw new Error(`格子里没有 w:tcPr，插不进边框：${cellText(chunk)}`);
  const body = pr[1] ?? '';
  let at = body.length;
  for (const tag of AFTER_BORDERS) {
    const i = body.indexOf(`<${tag}`);
    if (i >= 0 && i < at) at = i;
  }
  const patched = `<w:tcPr>${body.slice(0, at)}${xml}${body.slice(at)}</w:tcPr>`;
  return chunk.replace(pr[0], patched);
}

/**
 * 按「格子文字 → 边框」改写 docx 里的 `word/document.xml`，原地存回去。
 * 返回打中的补丁数；有 key 没打中就抛 —— 静默漏掉一条，量出来的结论会指向别的规则。
 */
export async function patchCellBorders(
  docxPath: string,
  patches: ReadonlyMap<string, RawCellBorders>,
): Promise<number> {
  if (patches.size === 0) return 0;
  const zip = unzipSync(new Uint8Array(await readFile(docxPath)));
  const partName = 'word/document.xml';
  const part = zip[partName];
  if (part === undefined) throw new Error(`${docxPath} 里没有 ${partName}`);

  const xml = new TextDecoder().decode(part);
  const ranges = cellRanges(xml);
  const hits = new Map<string, number>();
  let out = '';
  let at = 0;
  for (const { start, end } of ranges) {
    if (start < at) continue; // 嵌套表格的外层格子已经整段抄过了
    const chunk = xml.slice(start, end);
    const borders = patches.get(cellText(chunk));
    out += xml.slice(at, start);
    if (borders === undefined) {
      out += chunk;
    } else {
      const key = cellText(chunk);
      hits.set(key, (hits.get(key) ?? 0) + 1);
      out += applyToCell(chunk, borders);
    }
    at = end;
  }
  out += xml.slice(at);

  const missed = [...patches.keys()].filter((k) => !hits.has(k));
  if (missed.length > 0) throw new Error(`补丁没打中这些格子：${missed.join('、')}`);
  const dup = [...hits].filter(([, n]) => n > 1).map(([k]) => k);
  if (dup.length > 0) throw new Error(`这些文字在文档里不止一格，改哪个都不对：${dup.join('、')}`);

  zip[partName] = new TextEncoder().encode(out);
  // 固定时间戳：docx 要入库，用「现在」会让每次重建的二进制都不同。
  // 不能填 0 —— zip 的时间字段只表示得了 1980–2099，fflate 直接抛
  await writeFile(docxPath, zipSync(zip, { mtime: new Date('2020-01-01T00:00:00Z') }));
  return hits.size;
}
