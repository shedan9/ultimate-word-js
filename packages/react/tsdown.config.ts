import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['esm'],
  dts: true,
  clean: true,
  // react 是 peer；ultimate-word 与内部包各走各的产物
  external: [/^@uw\//, 'ultimate-word', /^react/],
});
