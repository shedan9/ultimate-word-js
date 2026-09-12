import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['esm'],
  dts: true,
  clean: true,
  // 内部包走各自的产物；fontkit 只在 `fonts.register()` 里动态 import，不进主 chunk
  external: [/^@uw\//],
});
