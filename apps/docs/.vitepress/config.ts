import { posix } from 'node:path';
import { defineConfig } from 'vitepress';

/** 仓库地址 */
const REPO = 'https://github.com/shedan9/ultimate-word-js';

/**
 * 按 GitHub 的规则生成标题锚点。docs/*.md 是照 GitHub 渲染写的，互相引用时手写的锚点
 * （`architecture.md#41-一个重要推论缩放永不触发重排`）是 GitHub 那一套 —— VitePress 默认
 * 把标点换成 `-`，同一个标题会得出 `4-1-一个重要推论-缩放永不触发重排`，站内跳转全部落空。
 */
function githubSlug(text: string): string {
  return text
    .trim()
    .toLowerCase()
    .replace(/[^\p{L}\p{M}\p{N}\p{Pc} -]/gu, '')
    .replace(/ /g, '-');
}

/** 原文里指向源码的相对链接（`../packages/…`，相对 docs/）在站点上没有对应页面，改指 GitHub */
function rewriteSourceLink(href: string): string | undefined {
  if (/^[a-z]+:|^#/i.test(href)) return undefined;
  const path = posix.normalize(href);
  if (path.startsWith('../packages/') || path.startsWith('../apps/'))
    return `${REPO}/blob/main/${path.slice(3)}`;
  return undefined;
}

export default defineConfig({
  lang: 'zh-CN',
  title: 'ultimate-word',
  description: '自研布局引擎的 Word（OOXML）在线预览 / 编辑库',
  cleanUrls: true,
  lastUpdated: false,
  // `docs/*.md` 通过 `<!--@include-->` 引进来，站点本身只放外壳页面
  srcExclude: ['README.md'],
  themeConfig: {
    nav: [
      { text: '在线预览', link: '/demo' },
      { text: 'API', link: '/api' },
      { text: '架构', link: '/architecture' },
      { text: '开发计划', link: '/DEVELOPMENT-PLAN' },
    ],
    outline: { level: [2, 3], label: '本页目录' },
    socialLinks: [{ icon: 'github', link: REPO }],
    search: {
      provider: 'local',
      options: {
        translations: {
          button: { buttonText: '搜索', buttonAriaLabel: '搜索' },
          modal: {
            noResultsText: '没有找到',
            resetButtonTitle: '清空',
            footer: { selectText: '打开', navigateText: '切换', closeText: '关闭' },
          },
        },
      },
    },
    docFooter: { prev: false, next: false },
    darkModeSwitchLabel: '外观',
    returnToTopLabel: '回到顶部',
    sidebarMenuLabel: '菜单',
  },
  markdown: {
    anchor: { slugify: githubSlug },
    // 站点页面与 docs/ 里的文件同名同层，原文的 `./architecture.md` 照原样就是对的；
    // 默认的改写会把它们变成 `./../../docs/architecture.md`，全成死链
    include: { rebaseRelativeUrls: false },
    config(md) {
      // ```mermaid 交给客户端组件画：mermaid 体积大且依赖 DOM，不进 SSR
      const fence = md.renderer.rules.fence;
      md.renderer.rules.fence = (tokens, idx, options, env, self) => {
        const token = tokens[idx];
        if (token?.info.trim() === 'mermaid') {
          return `<ClientOnly><MermaidDiagram code="${encodeURIComponent(token.content)}" /></ClientOnly>`;
        }
        return fence ? fence(tokens, idx, options, env, self) : self.renderToken(tokens, idx, options);
      };
      md.core.ruler.after('inline', 'uw-source-links', (state) => {
        for (const block of state.tokens) {
          for (const token of block.children ?? []) {
            if (token.type !== 'link_open') continue;
            const href = token.attrGet('href');
            const rewritten = href === null ? undefined : rewriteSourceLink(href);
            if (rewritten !== undefined) token.attrSet('href', rewritten);
          }
        }
      });
    },
  },
  vite: {
    server: { port: 5274 },
    // 与调试台同理：workspace 包的 exports 直接指向 src/*.ts，交给 Vite 现场编译
    optimizeDeps: {
      exclude: [
        '@uw/core',
        '@uw/fonts',
        '@uw/layout',
        '@uw/model',
        '@uw/ooxml',
        '@uw/render-dom',
        '@uw/serialize',
        '@uw/view',
        'ultimate-word',
      ],
    },
  },
});
