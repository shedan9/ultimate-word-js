# ultimate-word-js

自研布局引擎的 Word（OOXML）在线预览 / 编辑库。保真度由自己的引擎决定，不依赖浏览器排版。

定位：**中文公文 / 周报 / 总结 / 报告类文档的高保真引擎**，不是通用 Word 引擎。

| 文档 | 回答的问题 |
|---|---|
| [架构设计](./docs/architecture.md) | 代码怎么切、数据怎么流、坐标怎么管 |
| [API 设计](./docs/api.md) | 对外长什么样、为什么这样设计 |
| [开发计划](./DEVELOPMENT-PLAN.md) | 做什么、什么顺序、每阶段的完成判据 |

## 当前进度

Phase 0（地基与验证性穿刺）—— **穿刺已通过**。

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
```

## 目录

```
packages/
  core/          @uw/core        单位、几何、错误、事件、日志、类型     ← units.ts 已可用
  ooxml/         @uw/ooxml       OPC(zip) + XML → 原始 OOXML 树         ← Phase 1
  model/         @uw/model       文档模型 + 样式级联 + 编号 + 事务      ← Phase 1
  fonts/         @uw/fonts       字体表解析、度量、替换表、缓存         ← 行高规则已可用
  layout/        @uw/layout      布局引擎                               ← Phase 2
  render-dom/    @uw/render-dom  绝对定位 DOM 渲染器                    ← Phase 2
apps/
  playground/    调试台（Vite 8）
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
