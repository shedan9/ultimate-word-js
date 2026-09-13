import { defineConfig } from 'vite';

export default defineConfig({
  server: { port: 5273, open: false },
  // workspace 包的 exports 直接指向 src/*.ts，交给 Vite 现场编译，
  // 免得每改一行库代码就要先 build
  optimizeDeps: {
    // react 是 CJS，三个入口要一起预打包，否则 @uw/react（源码直读）里的 react-dom 与
    // 页面里的 react-dom/client 会拿到两份不同的 interop 产物
    include: ['react', 'react-dom', 'react-dom/client', 'react/jsx-runtime'],
    exclude: [
      '@uw/core',
      '@uw/fonts',
      '@uw/layout',
      '@uw/model',
      '@uw/ooxml',
      '@uw/react',
      '@uw/render-dom',
      '@uw/view',
      'ultimate-word',
    ],
  },
});
