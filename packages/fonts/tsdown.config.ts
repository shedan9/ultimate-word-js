import { defineConfig } from 'tsdown';

export default defineConfig({
  // decode 与 load-node 各自成入口：主入口不依赖 fontkit，
  // 只带度量包的部署不必把它打进包里
  entry: ['src/index.ts', 'src/decode.ts', 'src/load-node.ts'],
  format: ['esm'],
  dts: true,
  clean: true,
  // workspace 内部互相引用走源码，不打进产物
  external: [/^@uw\//],
});
