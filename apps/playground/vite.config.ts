import { defineConfig } from 'vite';

export default defineConfig({
  server: { port: 5273, open: false },
  // workspace 包的 exports 直接指向 src/*.ts，交给 Vite 现场编译，
  // 免得每改一行库代码就要先 build
  optimizeDeps: {
    exclude: ['@uw/core', '@uw/fonts', '@uw/layout', '@uw/model', '@uw/ooxml', '@uw/render-dom'],
  },
});
