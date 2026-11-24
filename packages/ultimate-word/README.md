# ultimate-word

在浏览器中打开、编辑和导出中文公文、周报和报告类 DOCX。适用于支持 ESM 的前端项目。内置常用字体的度量数据，实际显示的字形取决于设备字体。

将安装包复制到你的项目并运行：

```sh
npm install ./ultimate-word-0.1.1.tgz
```

页面提供一个文档容器，例如 `<div id="document"></div>`。将用户选中的 `File` 传给以下函数：

```ts
import { UltimateWord } from 'ultimate-word';

export async function openDocument(file: File) {
  const doc = await UltimateWord.load(file);
  const view = doc.mount('#document', { mode: 'edit' });
  return {
    // 点击保存时调用，返回包含当前编辑结果的 DOCX Blob。
    save: () => doc.toDocx(),
    // 切换文档或卸载页面时调用，释放当前视图。
    close: () => view.dispose(),
  };
}
```

只读预览可省略 `mode`。使用 `doc.find('关键词')` 获取正文命中范围，`doc.replaceAll('原文', '新文')` 替换可编辑的命中，`doc.undo()` 和 `doc.redo()` 操作编辑历史。加载后的 `doc.diagnostics` 提供文档内容或字体降级提示。

支持范围以中文横排 DOCX 为主，复杂公式、SmartArt 和多栏排版仍有限制。
