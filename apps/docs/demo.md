---
layout: page
title: 在线预览
---

<div class="uwd-page">

# 在线预览

文档在你的浏览器里解析和排版，不会上传。选一份样本，或者把自己的 .docx 拖进来。

<ClientOnly>
  <DocxDemo />
</ClientOnly>

<p class="uwd-note">
断行与分页用的是随库的 Word 字体度量，与 Word 一致；字形由本机字体画。本机没有宋体、仿宋这些字体时，
字的样子会不同，但每个字的位置不变。样本来自保真度真值流水线，都与 Word 导出的 PDF 逐行比过。
</p>

</div>

<style>
.uwd-page {
  max-width: 1152px;
  margin: 0 auto;
  padding: 32px 24px 64px;
}
.uwd-page h1 {
  margin: 0 0 8px;
  font-size: 28px;
  font-weight: 600;
  line-height: 1.3;
}
.uwd-page > p:first-of-type {
  margin: 0 0 24px;
  color: var(--vp-c-text-2);
}
.uwd-note {
  max-width: 72ch;
  margin: 16px 0 0;
  color: var(--vp-c-text-3);
  font-size: 13px;
  line-height: 1.7;
}
@media (max-width: 640px) {
  .uwd-page {
    padding: 20px 16px 48px;
  }
}
</style>
