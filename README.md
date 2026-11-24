# ultimate-word-js

自研布局引擎的 Word（OOXML）在线预览 / 编辑库。保真度由自己的引擎决定，不依赖浏览器排版。

定位：**中文公文 / 周报 / 总结 / 报告类文档的高保真引擎**，不是通用 Word 引擎。

| 文档 | 回答的问题 |
|---|---|
| [架构设计](./docs/architecture.md) | 代码怎么切、数据怎么流、坐标怎么管 |
| [API 设计](./docs/api.md) | 对外长什么样、为什么这样设计 |
| [开发计划](docs/DEVELOPMENT-PLAN.md) | 做什么、什么顺序、每阶段的完成判据 |

## 当前进度

`ultimate-word` 0.1.0 提供浏览器端 DOCX 分页预览、正文编辑、查找替换和 DOCX 导出。可以直接运行调试台，也可以把门面库安装到自己的前端项目。

```sh
pnpm install --frozen-lockfile
pnpm --filter @uw/playground dev   # 打开终端显示的网址，选择一份 .docx
pnpm pack:lib                      # 在仓库根目录生成 dist/ultimate-word-0.1.0.tgz
```

调试台可用仓库内的 `apps/fidelity/fixtures/gongwen-01.docx` 试用。勾选「编辑」后修改正文，点击「导出 docx」下载副本。集成方式见[安装包使用说明](./packages/ultimate-word/README.md)。

## 环境

Node 24（`.node-version`，fnm `--use-on-cd` 自动切换）+ pnpm 11（corepack，版本锁在 `packageManager` 字段）+
TypeScript 7 + Vite 8。工具脚本（`apps/fidelity`）直接用 Node 原生类型剥离跑 `.ts`，不经 tsx。
真值流水线还需要 **Windows + 已安装的 Word**（走 COM）；其余部分跨平台。

```bash
pnpm install
pnpm turbo run typecheck test    # 全量检查
pnpm truth                       # 重新生成保真度真值（需要 Word）
pnpm --filter @uw/fidelity spike # 跑 Phase 0 行高穿刺
pnpm --filter @uw/playground dev # 调试台
pnpm --filter @uw/docs dev       # 文档站 + 在线预览（http://localhost:5274）
```

## 目录

```
packages/
  ultimate-word/  对外的 ESM 门面与类型声明
  react/          React 视图与 hooks
  core/           单位、几何与诊断
  ooxml/          OPC 与 XML 解析
  model/          文档模型、样式级联与编辑事务
  fonts/          字体度量与替换
  layout/         断行、分页与表格布局
  render-dom/     页面绘制
  view/           视口、选区与输入
  serialize/      DOCX 回写
apps/
  playground/    调试台（Vite 8）
  docs/          文档站（VitePress）：docs/*.md + 在线预览 demo
  fidelity/      保真度真值流水线 + Phase 0 穿刺（见 apps/fidelity/README.md）
```

依赖方向严格单向：`core ← ooxml ← model ← layout ← render-* ← view ← editor`。
`layout` 不得 import 任何 DOM API。

## Phase 0 穿刺结论

**Word 的单倍行距不是 CSS 的 `line-height: normal`，也不是「1.3 × 字号」，
而是字体 win 度量跨度按脚本分两条路：**

| 行的内容 | 行高 |
|---|---|
| 含东亚文字 | `(usWinAscent + usWinDescent) × 1.3 × 字号 / unitsPerEm`，**不加**外部行距 |
| 纯拉丁文字 | `(usWinAscent + usWinDescent + GDI 外部行距) × 字号 / unitsPerEm` |

GDI 外部行距 = `max(0, hhea.lineGap - (win 跨度 - hhea 跨度))`。

13 个实测样本（仿宋 / 宋体 / 黑体 / 楷体 / 微软雅黑 / 等线 / Times New Roman / Arial），
**最大误差 0.132 pt**，其中约 0.1pt 是 Word 导出 PDF 时的坐标取整 —— 远低于 Phase 0 的 1pt 判据，
也低于 L3 断言的 0.5pt。实现在 `@uw/fonts` 的 `lineMetrics()`，回归测试在 `metrics.test.ts`。

> 分辨「1.3 系数」与「固定 1.3 倍字号」两种假设，靠的是 win 跨度差异极大的两款字体：
> 微软雅黑 1.3198 em（实测行高 1.71 em）与等线 1.0420 em（实测 1.35 em）。
> 只测宋体家族分不开 —— 它们的 unitsPerEm 是 256，win 跨度恰好是 1.0 em。

**未决**：这 30% 的额外行距在基线上下如何分配（决定行内基线的确切位置）。
Phase 2 之前必须再做一次「首行基线到版心顶」的穿刺定下来。

## 真值体系

开发机有 Word，所以真值是**坐标级**的，不是截图肉眼比对：
Word COM 导出 PDF → pdf.js 抽每个文本片段的 transform → `fixtures/*.truth.json`（入库）。
CI 上没有 Word，真值文件与 fixture docx 一并提交，本地跑 `pnpm truth` 重新生成。
详见 [apps/fidelity/README.md](./apps/fidelity/README.md)。
