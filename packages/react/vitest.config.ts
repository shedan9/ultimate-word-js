import { defineConfig } from 'vitest/config';

// 组件要真的挂进 DOM 才能验「挂 / 卸 / 换锚点」的接线，jsdom 只给这个包用
export default defineConfig({
  test: { environment: 'jsdom' },
});
