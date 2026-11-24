import { DOCX_MIME, UltimateWord, type UwView } from 'ultimate-word';

// 在仓库外执行，确保任何遗漏的工作区依赖都会在安装、类型检查或运行时暴露。
export async function verifyInstalled(file: Uint8Array, container?: HTMLElement): Promise<void> {
  const original = '各有关单位';
  const revised = '各有关部门';
  const doc = await UltimateWord.load(file);
  if (doc.pageCount !== 1 || doc.find(original).length !== 1) throw new Error('公文加载或分页失败');

  let view: UwView | undefined;
  try {
    if (container) {
      view = doc.mount(container, { mode: 'edit', virtualize: false });
      const range = doc.find(original)[0];
      if (!range || !view.rectsOf(range).length) throw new Error('正文没有可见坐标');
      view.select?.(range);
      if (!view.selection) throw new Error('编辑视图无法选择文字');
    }
    if (doc.replaceAll(original, revised).replaced !== 1) throw new Error('正文修改失败');
    doc.undo();
    if (doc.find(original).length !== 1) throw new Error('撤销未恢复原文');
    doc.redo();
    const blob = await doc.toDocx();
    if (blob.type !== DOCX_MIME) throw new Error('导出格式错误');
    const copy = await UltimateWord.load(blob);
    if (copy.find(revised).length !== 1 || copy.find(original).length) throw new Error('导出未保存修改');
  } finally {
    view?.dispose();
  }
}
