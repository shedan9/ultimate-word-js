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
```

新增 fixture 有两种方式：

1. **写 spec**（推荐用于最小复现）：在 `fixtures/src/` 放一个 JSON，Word 会据此生成 docx。
   文档正文只存在于这个 UTF-8 JSON 里，不写进 `.ps1`，避免 PowerShell 主机编码把中文吃掉。
2. **直接丢 docx**：把真实公文放进 `fixtures/`，没有同名 spec 也能跑。

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
- `Paragraphs.Add()` 不带 Range 时锚在 selection 上（`Documents.Add()` 后 selection 还在偏移 0），
  会把第二段并进第一段 —— 显式在文末位置插段落标记
- `ExportAsFixedFormat` 的 `BitmapMissingFonts` 必须为 `$false`，否则字被转成位图，PDF 里抽不到文本
- pdf.js v6 的 `destroy()` 在 **loadingTask** 上，不在 document proxy 上
- 字体真实名要先跑 `getOperatorList()` 填充 `commonObjs`，再用 `TextItem.fontName` 回查；
  v6 起 `commonObjs` 内部存储是私有字段，枚举不了，只能按 key 取
