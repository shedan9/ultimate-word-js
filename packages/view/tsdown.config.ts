import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: ['src/index.ts', 'src/dom.ts'],
  format: ['esm'],
  dts: true,
  clean: true,
  external: [/^@uw\//],
});
