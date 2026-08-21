# @uw/fidelity —— 保真度真值流水线

用本机安装的 Word 生成**坐标级真值**，而不是靠肉眼比截图。这是整个项目最高杠杆的基础设施：
没有真值，后面每一步布局代码都是在猜。

## 链路

```
fixtures/src/<name>.json  --Word COM-->  fixtures/<name>.docx
fixtures/<name>.docx      --Word COM-->  out/<name>.pdf + out/<name>.wordmeta.json
out/<name>.pdf            --pdf.js  -->  fixtures/<name>.truth.json   ← 入库
```

CI 上没有 Word，所以 **truth.json 与 docx 提交进仓库，PDF 不提交**（`out/` 已在 .gitignore）。
本机改了 fixture 就跑一次 `pnpm truth` 重新生成。

## 用法

```bash
pnpm truth                 # 只重算过期的 fixture
pnpm truth gongwen-01      # 指定 fixture（可多个）
pnpm truth --force         # 全部重跑
node src/extract-truth.ts out/gongwen-01.pdf   # 只抽某个已有 PDF，结果打到 stdout

pnpm spike                 # Phase 0 行高穿刺：单倍行距行高对不对
pnpm spike:baseline        # 基线穿刺：基线在行高里的位置对不对
pnpm spike:punct           # 标点挤压穿刺：什么时候压、压多少
pnpm spike:compress        # 临时挤压穿刺：塞不下时肯挤多少才换行
```

这四个 spike 脚本是**标定工具**，不是单测：它们从真值反推系数、打出残差表，
并在「最优假设不是代码里实现的那个」时以退出码 1 失败。跨平台的回归由单测兜着 ——
`@uw/fonts` 的 `metrics.test.ts` / `metrics-pack.test.ts` 与 `@uw/layout` 的
`items.test.ts` 里的期望值，就是这些脚本打出来的实测值。

各自钉死了什么：

| 脚本 | fixture | 结论 |
|---|---|---|
| `spike` | `spike-lineheight-01/02` | 单倍行距行高：东亚 win 跨度 × 1.3、拉丁 win 跨度 + 外部行距 |
| `spike:baseline` | `spike-baseline-01/02/03` | 基线位置：核心盒在行高里居中；网格吸附在行距倍数之前；东亚行的行盒只由东亚字体定；空段落走 ascii 桶 |
| `spike:punct` | `spike-punct-01` | 孤立标点不压，相邻标点固定压 0.5 em |
| `spike:compress` | `spike-compress-01/02` | 临时挤压只在两端对齐的行里发生；一个标点最多让 0.48 em；挤到什么程度就宁可换行（`挤压量 × 字距数 ≤ 30.6 × 标点数 × 拉伸量`） |

新增 fixture 有两种方式：

1. **写 spec**（推荐用于最小复现）：在 `fixtures/src/` 放一个 JSON，Word 会据此生成 docx。
   文档正文只存在于这个 UTF-8 JSON 里，不写进 `.ps1`，避免 PowerShell 主机编码把中文吃掉。
2. **直接丢 docx**：把真实公文放进 `fixtures/`，没有同名 spec 也能跑。

### spec 认的字段

页面：`widthMm` / `heightMm` / `marginMm`，以及可选的 `grid`
（`{ linesPage }` = 只吸基线的行网格；再给 `charsLine` 就升级成连汉字也吸到列上的字符网格）。
**没写 `grid` 就显式关掉网格** —— 中文版 Word 的 Normal 模板默认是开着的。

段落：`text`（可以是空串，用来量空段落的行高）、`fontEA` / `fontLatin` / `sizePt` / `bold` /
`align` / `firstLineChars` / `leftIndentPt` / `rightIndentPt` / `spaceBeforePt` / `spaceAfterPt`，行距二选一
（`lineSpacingPt` = 固定值，`lineSpacingMultiple` = 倍数），以及
`pageBreakBefore`（把这一段顶到新页的最上面，量「首行基线 − 版心顶」靠它）
与 `snapToGrid: false`（这一段不吸网格）。

缩进有两套单位，别混：`firstLineChars` 是**字符**（1/100 字），`leftIndentPt` / `rightIndentPt`
是**点**。要把可用宽度一格格调窄来找某个阈值，用 pt 的那套 —— 字符单位既量化到字号的 1/100，
又依赖「一个字符多宽」这个待标定的量（`spike-compress-01/02` 就是踩着这一点设计的）。

一段只有**一个字号**：`w:rPr` 是挂在整段 range 上的。要造「一行里两个字号」的样本
得先给 `make-fixture.ps1` 加 run 级支持，`@uw/fonts` 的 `composeBaseline()` 正等着这个。

## truth.json 读法

- 单位 **pt**，原点**页面左上角、y 向下**（PDF 原生是左下原点，抽取时已翻转）
- `pages[].items[]`：一个 PDF show-text 片段 ≈「同字体的连续字符段」，`y` 是**基线**
- `pages[].lines[]`：按基线容差 0.6pt 聚合出的行，`first` / `last` 是行首末**码点** ——
  L1（每页首末行）、L2（每行断行点）级断言直接用它
- `fonts`：pdf.js 从字体表读出的归一化 `ascent` / `descent`（已除以 unitsPerEm）。
  这正是 Word 算行高的输入，Phase 0 的行高验证拿它对
- `sections` / `wordPageCount`：来自 Word COM 自述，不是从 PDF 反推的，用于交叉校验

输出刻意做成**确定性**的：不写时间戳，浮点固定 3 位小数，字体子集前缀（`BCDEEE+`）已剥掉 ——
否则每次重新生成，diff 全是噪声。

## 断言分级（随阶段收紧）

| 级别 | 断言内容 | 目标阶段 |
|---|---|---|
| L0 | 总页数一致 | Phase 3 |
| L1 | 每页首行 / 末行文本一致 | Phase 3 |
| L2 | 每行断行点一致（行首末字符） | Phase 2 |
| L3 | 每行基线 y 误差 < 0.5pt | Phase 2 |
| L4 | 每个 run 片段起始 x 误差 < 0.5pt | Phase 4 |

## 环境要求

Windows + 已安装 Word（走 COM）。

脚本走 `pwsh` 优先、Windows PowerShell 5.1 回落两条路，**两条都验过**：
开发机是 pwsh 7.6.4（日常跑的就是它），5.1（5.1.26100）跑同一个 spec 产出的真值与
入库版本逐行文本一致、基线差 0.000pt。所以 `.ps1` 里不用 PS7 专属语法，
中文也一律留在 UTF-8 的 spec JSON 里由 `[System.IO.File]::ReadAllText` 显式读，
不依赖主机默认编码 —— 这样两个版本的行为才是一样的。

## 已知坑（踩过的）

- `Range.Text = ...` 会连**段落标记**一起替换，把所有段落并成一段 —— 必须用 `InsertBefore`
- **段落格式会被后插入的段落继承**：`InsertParagraphAfter` 复制上一段的 `ParagraphFormat`，
  所以 `PageBreakBefore` 这类开关必须**无条件赋值**（spec 里没写就赋 `$false`），
  写成 `if (spec 里有) { 设 }` 会让第 4 段的分页符泄漏到后面每一段 —— 17 段各占一页，
  第一次跑基线穿刺就是这么废掉的。脚本里其余 `$fmt.*` 一直是无条件赋值，正是这个原因
- `Paragraphs.Add()` 不带 Range 时锚在 selection 上（`Documents.Add()` 后 selection 还在偏移 0），
  会把第二段并进第一段 —— 显式在文末位置插段落标记
- `ExportAsFixedFormat` 的 `BitmapMissingFonts` 必须为 `$false`，否则字被转成位图，PDF 里抽不到文本
- pdf.js v6 的 `destroy()` 在 **loadingTask** 上，不在 document proxy 上
- 字体真实名要先跑 `getOperatorList()` 填充 `commonObjs`，再用 `TextItem.fontName` 回查；
  v6 起 `commonObjs` 内部存储是私有字段，枚举不了，只能按 key 取
