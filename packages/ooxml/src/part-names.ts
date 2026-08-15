/**
 * 部件名与路径解析。
 *
 * OPC 的部件名是**带前导斜杠的绝对路径**（`/word/document.xml`），
 * 而 zip 条目名不带斜杠（`word/document.xml`），关系里的 Target 又可能是相对路径
 * （`styles.xml`、`../docProps/app.xml`）。三种形态混着用是这一层最容易出 bug 的地方，
 * 所以规矩定死：**内部一律用带斜杠的绝对部件名**，只在读 zip 时脱掉斜杠。
 */

/** zip 条目名 → 部件名 */
export function toPartName(zipEntryName: string): string {
  return zipEntryName.startsWith('/') ? zipEntryName : `/${zipEntryName}`;
}

/** 部件名 → zip 条目名 */
export function toZipEntryName(partName: string): string {
  return partName.startsWith('/') ? partName.slice(1) : partName;
}

/** `/word/document.xml` → `/word`；根部件返回空串 */
export function dirNameOf(partName: string): string {
  const i = partName.lastIndexOf('/');
  return i <= 0 ? '' : partName.slice(0, i);
}

/**
 * 关系目标 → 绝对部件名。
 *
 * 基准是**源部件所在目录**，不是 `.rels` 文件所在目录 —— `/word/_rels/document.xml.rels`
 * 里写 `styles.xml` 指的是 `/word/styles.xml` 而不是 `/word/_rels/styles.xml`。
 * 这个坑踩过一次就忘不掉：所有样式表都会 404。
 */
export function resolveTarget(sourcePartName: string, target: string): string {
  if (target.startsWith('/')) return normalizePath(target);
  const base = dirNameOf(sourcePartName);
  return normalizePath(`${base}/${target}`);
}

/** 折叠 `.` 与 `..`，输出带前导斜杠的规范路径 */
function normalizePath(path: string): string {
  const out: string[] = [];
  for (const seg of path.split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') out.pop();
    else out.push(seg);
  }
  return `/${out.join('/')}`;
}

/** 部件名 → 它的关系部件名。包级关系传空串，得到 `/_rels/.rels` */
export function relsPartNameOf(partName: string): string {
  if (partName === '' || partName === '/') return '/_rels/.rels';
  const dir = dirNameOf(partName);
  const base = partName.slice(dir.length + 1);
  return `${dir}/_rels/${base}.rels`;
}

/** 扩展名，小写，不含点。没有扩展名返回空串 */
export function extensionOf(partName: string): string {
  const i = partName.lastIndexOf('.');
  const slash = partName.lastIndexOf('/');
  return i > slash ? partName.slice(i + 1).toLowerCase() : '';
}
