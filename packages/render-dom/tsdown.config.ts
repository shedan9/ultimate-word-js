import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: ['src/index.ts', 'src/dom.ts'],
  format: ['esm'],
  dts: true,
  clean: true,
  // workspace 内部互相引用走源码，不打进产物
  external: [/^@uw\//],
});
