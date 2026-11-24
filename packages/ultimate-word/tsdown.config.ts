import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['esm'],
  tsconfig: '../tsconfig.dist.json',
  dts: true,
  clean: true,
  fixedExtension: true,
  platform: 'browser',
  // 分发包收进内部实现和公开签名，第三方运行时依赖交给宿主打包器解析。
  deps: {
    alwaysBundle: [/^@uw\//],
    onlyBundle: [],
  },
});
