# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 这个项目是什么

自研布局引擎的 Word（OOXML）在线预览 / 编辑库，定位是**中文公文 / 周报 / 报告类文档的高保真引擎**。
保真度由自己算的排版决定，**不依赖浏览器排版** —— 所以「用 CSS 让它看起来差不多」永远不是正确答案。

当前进度：Phase 0 已完成（地基 + 行高穿刺 + CI），Phase 1 的**解析链已完整**，
Phase 2 已排到**行盒的门口** —— 水平方向（断行 + 行内几何）与行高总量都做完了，就差基线。
Phase 5 的**列表编号**也已经从 `numbering.xml` 一路通到首行几何（它不依赖基线，所以能先做）。
Phase 4 的**表格**也排到了同一个门口：属性 + 级联（含 `w:tblStylePr` 条件格式）在 model 层，
列宽 + 每格的 x 与可用宽 + 格内段落在 layout 层，**没有 y**。
真实实现：`@uw/core`（单位 / 错误 / 诊断）、`@uw/ooxml`（OPC 容器 + XML 树）、
`@uw/model`（样式级联 + 主题字体 + 正文节点树 + 分节 + 设置 + 字体表 + 制表位 + **编号（解析 + 计数器 + 编号文字 + 接进级联）** + **表格（属性 + 级联 + 条件格式）**）、
`@uw/fonts`（行高规则 + 脚本分桶 + 度量包 + 注册表 + `TextMeasurer`）、
`@uw/layout`（item 流 + 断行 + 缩进 / 对齐 / 制表位 / 列表编号 + 行高与网格吸附 + **表格列宽与格内几何**）。
`@uw/render-dom` 仍是占位 —— 没有 y 画不了，等下面那个穿刺。

一份 docx 的入口是 `loadDocument(pkg, sink)`（`packages/model/src/load.ts`），产出
`body`（直接格式，可编辑）、`resolved`（级联完的纯数据，给布局）、`cascade`（上下文，**不可**过 Worker 边界）、
`fonts`、`numbering`。`resolveBody()` 那一步就是 Worker 边界 —— `StyleSheet` 带方法，级联必须在过界前做完；
它同时是**编号计数器**跑的地方（编号「第几」只有按文档顺序走一遍才知道），结果落在
`ResolvedParaProps.numbering.label`（编号文字 + 它自己的字符属性 + `w:suff`）。
部件一律**按关系类型找**（`RelType.*`），不按 `word/styles.xml` 这种路径惯例猜。

**卡在 Windows 的两件事**：
① 东亚行高里那 30% 额外行距在基线上下如何分配，需要「首行基线到版心顶」的穿刺来定（要 Word COM）——
它挡的是**行盒装配及其之后的一切**（基线、y、分页、DOM 渲染）；
② A/C 类中文字体的**度量包抽取**（要 `C:/Windows/Fonts`）—— 没有它，Mac 上只能走三级降级的
第③级等宽近似，于是**坐标级真值断言（L2–L4）跑不了**，断行算法的正确性目前只由合成字体的单测保证。
**不挡**解析、样式级联、分桶、度量、断行、行内几何 —— 这些在 Mac 上都能做完，也确实做完了。

`@uw/layout` 的用法：`layoutParagraph(resolvedParagraph, { measurer, contentWidth, settings, docGrid })`
→ `ParagraphLayout`（每行的 x / 逐字 x / 行高 / 渲染片段，**没有 y**）；
`layoutTable(resolvedTable, { …, availWidth })` → `TableLayout`（列宽 / 每格的 x 与可用宽 /
格内段落，同样**没有 y**）。表格的列宽直接取 `w:tblGrid` —— **Word 存盘时已经把 autofit
算完的结果写在那儿了**，照着用就与 Word 一致，这也是「完整 autofit 算法」能列为非目标的原因。
未标定的常数一律集中在 `packages/layout/src/uncalibrated.ts`，每条都写了「拿什么样本能钉死」——
布局里出现别处的魔法数字视为 bug，散落的数字会被后人当成实测结论。

`@uw/fonts` 的用法：`FontRegistry` 收字体（`fontkitSource` 一级 / `metricsPackSource` 二级；
字节走 `@uw/fonts/decode` 的 `fontSourceFromBytes()`，文件走 `@uw/fonts/node` 的 `fileSource()`
—— **主入口刻意不依赖 fontkit**，只带度量包的部署不该被迫打包它），
`createTextMeasurer(registry, { candidates, diagnostics })` 产出给 layout 注入的 `TextMeasurer`。
`candidates` 那个回调就是把 model 的 `fontNameCandidates()` 接进来的地方 —— fonts **不认识** model，
依赖方向不许反过来。文字先过 `splitFontRuns(text, fonts)` 切成「同字体同脚本」的段，再逐段量。

级联里**写明的洞**（别以为已经做了）：toggle 属性在样式层之间的 XOR 语义（§17.7.3，现按「后者覆盖」处理），
`cascade.ts` 末尾有注释说明补的办法。

编号那一层已经接上，三处容易搞反：
① `w:lvl/w:pPr`（缩进）铺到**整个段落**，`w:lvl/w:rPr` 只作用于**编号文字**（铺到正文上会「整段变 Symbol」）；
② 计数按 **numId** 分家而不是 abstractNumId —— 「重新开始编号」正是靠两个 num 指同一个 abstractNum 实现的；
③ 布局里编号是一段**没有 run 的 item**（`numbering: true` 标记），命中测试与可选文本层必须跳过它。
未做：`w:lvlJc`（编号自身对齐）只带在数据上、布局忽略；中文读法的几处未标定见 `number-format.ts` 文件头。

表格那一层（`table-props.ts` / `parse-table-props.ts` / `cascade-table.ts`）也有四处容易搞反：
① 级联层序是「样式链自身属性 → 命中的条件格式（按 `CONDITIONAL_ORDER`）→ 直接格式」，
其中**行带排在列带之后、首末行排在首末列之后**（所以表头行会盖住首列的格式）；
② `w:tblLook` 是**开关**不是格式 —— 样式里定义了 `firstRow` 但 look 说不要，那份格式就不应用；
③ 表格样式的 `pPr` / `rPr` 铺给格内段落时排在**段落样式链之前**（走 `CascadeContext.tableStyleLayers`），
段落自己的样式要能盖掉表头行的加粗；
④ 单元格左右各 108 twips 的默认边距来自**默认表格样式**（`Normal Table`）而不是什么规范常数 ——
`w:tcMar` 缺席退到表级 `w:tblCellMar`，不是退到 0。
未标定：隔行带（`band1Horz` 那四种）的**序号算法**只有规范做依据，没有 Word 样本，见 `cascade-table.ts` 文件头。

字体名有个坑：中文版 Word 写的是「黑体」「等线」这种本地化名，磁盘上的字体叫 `SimHei` / `DengXian`。
桥在 `fontTable.xml` 的 `w:altName`，查找顺序用 `fontNameCandidates()`，别只按一个名字查。

三份必读文档，动手前按需查：

| 文档 | 回答 |
|---|---|
| [docs/architecture.md](docs/architecture.md) | 代码怎么切、数据怎么流、坐标怎么管、五条设计原则 |
| [docs/api.md](docs/api.md) | 对外 API 长什么样 |
| [docs/DEVELOPMENT-PLAN.md](docs/DEVELOPMENT-PLAN.md) | 阶段顺序、每阶段 DoD、非目标清单 |
| [apps/fidelity/README.md](apps/fidelity/README.md) | 真值流水线怎么用、truth.json 怎么读、踩过的坑 |

## 环境

Node 24（`.node-version` = 24.19.0，`engines.node >= 24.12.0`）+ pnpm 11 + TypeScript 7 + Vite 8。
非交互 shell 里 fnm 不注入，跑命令前先挂 PATH（用 24，v22 会触发 engine 警告）：

```shell
export PATH="$HOME/.local/share/fnm/node-versions/v24.19.0/installation/bin:$PATH"
```

工具脚本（`apps/fidelity`）**直接用 `node src/xxx.ts` 跑**，靠 Node 24 原生类型剥离，不经 tsx —— 别引入 tsx / ts-node。

## 常用命令

```bash
pnpm install
pnpm turbo run typecheck test        # 全量检查（跨平台，应当全绿）
pnpm lint                            # biome check .
pnpm lint:fix
pnpm --filter @uw/playground dev     # 调试台，:5273

# 单个包 / 单个测试文件 / 单个用例
pnpm --filter @uw/fonts run test
pnpm --filter @uw/fonts run test src/metrics.test.ts
pnpm --filter @uw/fonts run test src/metrics.test.ts -t "东亚"

# 真值流水线（仅 Windows + 已装 Word）
pnpm truth                           # 只重算过期 fixture
pnpm truth gongwen-01                # 指定 fixture
pnpm truth --force
pnpm --filter @uw/fidelity spike     # Phase 0 行高穿刺
```

真值的**生成**绑死 Windows（Word COM + `C:/Windows/Fonts`），真值的**消费**跨平台 ——
`fixtures/*.truth.json` 与 fixture docx 都已入库，PDF 不入库。在 Mac/Linux 上跑 `pnpm truth`
会被 `apps/fidelity/src/platform.ts` 的守卫拦下并以退出码 2 退出，这是设计如此，不是 bug。

## 架构约束（违反了后面搬不动）

1. **阶段之间只传纯数据。** 每个阶段的输出必须可结构化克隆：不含类实例方法、闭包、DOM 引用、反向指针。
   这一条同时买到 Worker 化、Rust/WASM 替换、golden file 回归三样能力。
2. **`@uw/layout` 不得 import 任何 DOM API。** 需要测量文本时用注入的 `TextMeasurer` 接口，实现由 `@uw/fonts` 给。
   `linebreak` / `measure` 的接口保持纯数据进出（码点数组 + 宽度数组 → 断点数组）。
3. **单位只有 twips，px 只在渲染出口。** 转换全部走 [`@uw/core/units.ts`](packages/core/src/units.ts)，
   布局过程中出现 px 视为 bug（逐行累加的浮点漂移会让最后一行溢出版心）。
   推论：缩放只作用在「布局空间 → 屏幕空间」这一步，**永不触发重排**。
4. **未识别的 XML 必须原样保留**，回写时吐回去 —— round-trip 安全是信誉底线。
5. **诊断 vs 异常不要混**：结构性错误（不是 zip、缺 `document.xml`）**抛**；
   内容问题（不认识的元素、basedOn 成环、字体缺失）**记 `Diagnostic` 并继续渲染其余部分**。
   用户要的是看到文档，不是看到白屏。

包依赖方向严格单向：`core ← ooxml ← model ← layout ← render-* ← view ← editor`，
`serialize ← model`。虚线以内（core / ooxml / model / fonts / layout / serialize）无 DOM 依赖，
这条线就是未来的 Worker 边界。

三个坐标空间、两个转换：`DocPosition{nodeId,offset}` ↔（`LayoutIndex`）↔ `LayoutPoint{page,x,y}`（twips）
↔（`ViewTransform`）↔ `ClientPoint`（px）。模型层永远不知道像素，渲染层永远不知道 `DocPosition`。

## 真值驱动

`apps/fidelity` 用 Word COM 导出 PDF、pdf.js 抽每个文本片段的 transform，产出**坐标级**真值
`fixtures/*.truth.json`（单位 pt，原点页面左上角、y 向下、`y` 是基线）。这不是测试工具，是架构的一部分：
它直接约束 `LayoutResult` 的数据形状 —— 必须能逐行、逐片段与真值 diff。

断言分级随阶段收紧：L0 页数 → L1 每页首末行文本 → L2 每行断行点 → L3 基线 y 误差 < 0.5pt → L4 片段起始 x 误差 < 0.5pt。

写布局代码时的判据是「与真值差多少 pt」，不是「看起来像不像」。改 `metrics.ts` 之类的标定结果前先看
`metrics.test.ts` —— 里面的期望值是从 Word 实测反推的，真值站在测试那一边。

## 已标定与未决

**已定死（Phase 0，13 个样本最大误差 0.132pt）** —— 实现在 [`packages/fonts/src/metrics.ts`](packages/fonts/src/metrics.ts)：

| 行的内容 | 单倍行距行高 |
|---|---|
| 含东亚文字 | `(usWinAscent + usWinDescent) × 1.3 × 字号 / unitsPerEm`，**不加**外部行距 |
| 纯拉丁文字 | `(usWinAscent + usWinDescent + GDI 外部行距) × 字号 / unitsPerEm` |

那个 1.3 是乘在**字体度量**上，不是「1.3 × 字号」；只测宋体家族（unitsPerEm=256、win 跨度恰好 1.0em）两种假设分不开。

**未决（阻塞 Phase 2 的行盒）**：东亚那 30% 额外行距在基线上下如何分配 —— 决定行内基线的确切位置，
现在一律记进 `lineGap`，只保证行高总量正确。**写行盒布局代码之前必须先补「首行基线到版心顶」的穿刺。**

行高之外还有两处未标定。一是 `splitFontRuns()` 的歧义字符集取的是 Unicode **EastAsianWidth = Ambiguous**
（`w:hint` 要回答的正是「这份文档算不算东亚环境」，两者同构），但 Word 的实际边界有没有偏差没有真值验证过 ——
上 Windows 时顺手做一份「① ※ ℃ Ⅰ 在 hint=eastAsia / default 下各占多宽」的样本就能钉死。
二是编号的三个样本：`w:lvlJc="right"` 的编号以哪条线对齐、编号宽过悬挂缩进时正文落在哪、
`chineseCounting` 与 `chineseCountingThousand` 在 105 / 1005 上各显示什么。
三是表格隔行带的序号：一份「6 行 3 列、开表头行 + 隔行带、`w:tblStyleRowBandSize=2`」的样本，
就能钉死「首行算不算进带」「带从第几条开始数」这两问 —— 带影响字重，字重影响宽度，最终影响断行。

另一个反复咬人的点：中文版 Word 的 Normal 模板**默认开着行网格**（linePitch 312 twips = 15.6pt），
基线会被吸到网格上、把字体度量差异整个盖掉。做度量实验必须显式关网格（`PageSetup.LayoutMode = wdLayoutModeDefault`）。

## 代码约定

- **中文注释**，且注释写「为什么这么定 / 违反会怎样」，不写「这是什么」。既有文件的密度就是基准。
- **相对 import 必须带 `.ts` 扩展名**（`allowImportingTsExtensions` + `rewriteRelativeImportExtensions`）。
- `erasableSyntaxOnly` 开着：不能用 enum、参数属性、namespace 等需要运行时代码生成的 TS 语法。
- `strict` + `exactOptionalPropertyTypes` + `noUncheckedIndexedAccess` + `verbatimModuleSyntax`（类型导入写 `import type`）。
- workspace 包的 `exports` 直接指向 `src/index.ts` —— **开发时不需要先 build**，Vite / vitest / node 现场编译。
  `tsdown` 只在产库时用，且 `external: [/^@uw\//]`，内部依赖不打进产物。
- Biome：单引号、分号、trailing comma、行宽 110、2 空格。`*.truth.json` 已排除在 lint 之外。
- 新包按现有包照抄结构：`package.json`（exports → src）+ `tsconfig.json`（extends base）+ `tsdown.config.ts`。
- `.ps1` 在工作区是 CRLF（`.gitattributes` 规定），且不用 PS7 专属语法 —— pwsh 7 与 Windows PowerShell 5.1 两条路都要能跑。
  中文正文一律放 UTF-8 的 spec JSON 里由 `[System.IO.File]::ReadAllText` 读，不写进 `.ps1`，免得被主机编码吃掉。

## 非目标（不要为它们预留扩展点）

紧密型/穿越型环绕、多栏排版、修订痕迹的**编辑**（只做显示）、OMML 数学公式排版、
VML / SmartArt / 图表（降级占位图）、`.doc` 二进制格式、完整 autofit 表格算法、RTL 与复杂文字。
需要时再改架构，比现在摆一堆空接口便宜。
