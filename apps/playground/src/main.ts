/**
 * 调试台。Phase 2 起会长成「可视化布局盒 + 逐帧排版」的样子；
 * 现在先只证明 workspace 包能在浏览器侧被正常引用与构建。
 */
import { mmToTwips, TWIP_PER_INCH, TWIP_PER_PT, twipsToPt, twipsToPx } from '@uw/core';

const a4WidthTwips = mmToTwips(210);
const rows: Array<[string, string]> = [
  [
    'A4 页宽',
    `${a4WidthTwips.toFixed(0)} twips = ${twipsToPt(a4WidthTwips).toFixed(2)} pt = ${twipsToPx(a4WidthTwips).toFixed(2)} px`,
  ],
  ['单位常量', `TWIP_PER_INCH=${TWIP_PER_INCH}，TWIP_PER_PT=${TWIP_PER_PT}`],
  ['下一步', 'Phase 1：OOXML 解析 + 文档模型 + 样式级联'],
];

const app = document.querySelector<HTMLElement>('#app');
if (app) {
  app.innerHTML = `
    <h1>ultimate-word 调试台</h1>
    <table>${rows.map(([k, v]) => `<tr><th align="left">${k}</th><td>${v}</td></tr>`).join('')}</table>
  `;
}
