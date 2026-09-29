<script setup lang="ts">
/**
 * ```mermaid 代码块的客户端渲染（config.ts 里把 fence 换成了这个组件）。
 * 跟随站点的亮 / 暗模式重画：mermaid 的配色是画的时候定死在 SVG 里的，切主题不重画就一直是旧的。
 */
import { useData } from 'vitepress';
import { onMounted, ref, watch } from 'vue';

const props = defineProps<{ code: string }>();
const { isDark } = useData();
const svg = ref('');
const failed = ref('');
let seq = 0;

async function render(): Promise<void> {
  const { default: mermaid } = await import('mermaid');
  // 节点尺寸是 mermaid 按它初始化时的字体量出来的；画出来的字却吃站点的字体，两边不一致就会裁字
  const fontFamily = getComputedStyle(document.body).fontFamily;
  mermaid.initialize({
    startOnLoad: false,
    theme: isDark.value ? 'dark' : 'default',
    securityLevel: 'strict',
    fontFamily,
    themeVariables: { fontFamily },
  });
  try {
    const id = `uw-mermaid-${Math.random().toString(36).slice(2)}-${seq++}`;
    svg.value = (await mermaid.render(id, decodeURIComponent(props.code))).svg;
    failed.value = '';
  } catch (err) {
    failed.value = err instanceof Error ? err.message : String(err);
  }
}

onMounted(render);
watch(isDark, render);
</script>

<template>
  <div class="uw-mermaid">
    <!-- biome-ignore lint/security/noDangerouslySetInnerHtml: mermaid 以 strict 模式输出的 SVG -->
    <div v-if="svg" v-html="svg" />
    <pre v-else-if="failed">{{ decodeURIComponent(code) }}</pre>
  </div>
</template>

<style>
.uw-mermaid {
  margin: 16px 0;
  overflow-x: auto;
}
/* 标签是 foreignObject 里的 <p>，.vp-doc p 的 28px 行高与上下外边距会把它撑出节点、被裁掉 */
.vp-doc .uw-mermaid p {
  margin: 0;
  line-height: 1.5;
}
.uw-mermaid svg {
  display: block;
  max-width: 100%;
  height: auto;
  margin: 0 auto;
}
</style>
