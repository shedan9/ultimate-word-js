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
pnpm spike:page            # 分页穿刺：孤行寡行 / keepNext / 页首段前间距
pnpm spike:header          # 页眉页脚穿刺：框摆在哪、怎么反过来挤版心
pnpm spike:image           # 图片穿刺：内嵌图在行盒里怎么摆、浮动图的参照框是哪个

pnpm preview                     # 全部 fixture → out/*.html，用眼睛看引擎画成什么样
pnpm preview gongwen-01 -- --truth --debug   # 叠真值基线（红虚线）+ 画版心与行盒
```

这七个 spike 脚本是**标定工具**，不是单测：它们从真值反推系数、打出残差表，
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
| `spike:page` | `spike-page-01/02` | 孤行寡行保底 2 行；段前间距落在页首不算；keepNext 的接缝要留出下一块「最少能放多少」 |
| `spike:header` | `spike-header-01/02/03` | 页眉框顶 = `w:header`；页脚量的是框**底**；页边距是最小值（页眉页脚长过它就把版心顶开） |
| `spike:image` | `spike-image-01/02` | 内嵌图占的高度 = 图高四舍五入到 1.5pt（坐在基线上的是这个**盒**，图在盒里靠上放）；文字的下伸留着；`w:position` 对图片起作用；浮动图八种参照框各是哪个（纵向的 inside/outside 镜像的是**上下页边距**、`character` 参照的是锚点**前一个**字） |

`spike:baseline` 现在跑**四份** fixture：04 补的是前三份漏掉的那一格 —— **固定值行距**
（`w:lineRule="exact"`）下基线 = 行高 × 0.8，与字体、字号都无关。它是被 `spike-page-01`
逼出来的：那份样本用固定行距 20pt，整页文字比「核心盒居中」的预测低 1.77pt。

`spike:page` / `spike:header` / `spike:image` 跟其余四个不同，它们**不反推系数**，而是把整台引擎跑一遍再与真值逐页对：
分页规则不是一个数，是三条互相纠缠的判断，单独反推任何一条都会被另一条污染。
做法是把 3 × 2 × 3 种组合排开，看哪一组能逐页复现 Word（当前：唯一满分 50/50 页）。
它们也是**不需要 Windows** 的 spike —— docx 与 truth.json 都入库了。

`spike:header` 的三份样本各占一格，缺一格就分不开：01 的页眉页脚**放得下**（正文一动不动）、
02 的**放不下**（该顶开正文、一页少排三行）、03 开「首页不同 + 奇偶页不同」并在页脚里放
**真的** `{ PAGE }` 域。8 种组合（页脚量顶 / 量底 × 挤版心的四种）逐页跑，唯一满分 12/12 页。
比对时页眉、正文、页脚的行是**混在一起按 y 排**再比的 —— 真值来自 PDF，
而 PDF 里没有「这是页眉」这回事。

`spike:image` 是唯一要用到真值里 `images[]` 那一路的脚本，另有两处与别的 spike 不同：

- **行比的是逐行增量，不是累加的绝对 y**。Word 自己的行位置带着 ±0.12pt 抖动
  （01 里纯文字参照行的行距在 15.48–15.62pt 之间跳），44 行图叠起来累到 1.5pt，
  早已越过 L3 —— 那是 Word 内部取整的锅，不是行盒规则的锅，而规则决定的恰好是增量
- **内嵌图比的是它相对本行基线的抬升**，浮动图才比纸坐标，理由同上

01 的图高排成三条阶梯：粗的 4→60pt 十三档、细的 30→36pt 步长 0.5pt、微的 30.0→31.5pt
步长 0.1pt。三条不是冗余 —— 粗阶梯只取偶数 pt 时，1.5pt 量化那条规则表现为
「h ≡ 4 (mod 6) 的那几档凭空多抬半磅」，看着像噪声，微阶梯才看得出是台阶。

## preview：用眼睛验收

`pnpm preview` 把 fixture 走完整条链（解包 → 级联 → 度量 → 断行 → 分页 → 画）再落盘成
`out/<name>.html`（**不入库**，与 PDF 同理）。它不需要 Word，也不需要 Windows。

`--truth` 会在每一页叠一层真值：Word 画每一行时的基线是**红虚线**，我们自己的基线是
**蓝实线**，重合就是对的。之所以能直接叠上去而不换算，是因为渲染器的 `viewBox` 单位
选的就是 pt，原点也同样是纸的左上角、y 向下 —— 与 `truth.json` 是同一套坐标系。

`--debug` 额外画出版心框（绿虚线）与每一行的行盒（蓝细框），排版跑偏时一眼能看出
是版心算错了还是行高算错了。

⚠️ 命令行末尾报的「N 行基线最大差 x pt」**不是保真度指标**：它按行序号硬配对，
只要有一行断得与 Word 不同，后面每一行都会错位一整行的高度，报出来的就是十几 pt 的假差值。
真正的判据在 `packages/layout/src/fixture.test.ts`（L2 / L3）与上面几个 spike 脚本里。
这个数只回答一个问题：「刚才那一改，有没有把某一页整体挪歪」。

新增 fixture 有两种方式：

1. **写 spec**（推荐用于最小复现）：在 `fixtures/src/` 放一个 JSON，Word 会据此生成 docx。
   `page` 里可以写 `headerDistMm` / `footerDistMm` / `differentFirstPage` / `differentOddEven`，
   顶层的 `headers` / `footers` 各接 `default` / `first` / `even` 三份段落列表；
   段落上写 `"field": "PAGE"` 会在它末尾插入一个**真的**域（不是打上去的数字）。
   文档正文只存在于这个 UTF-8 JSON 里，不写进 `.ps1`，避免 PowerShell 主机编码把中文吃掉。
2. **直接丢 docx**：把真实公文放进 `fixtures/`，没有同名 spec 也能跑。

### spec 认的字段

页面：`widthMm` / `heightMm` / `marginMm`，以及可选的 `grid`
（`{ linesPage }` = 只吸基线的行网格；再给 `charsLine` 就升级成连汉字也吸到列上的字符网格）。
**没写 `grid` 就显式关掉网格** —— 中文版 Word 的 Normal 模板默认是开着的。

段落还可以带 `images`：`{ file, widthPt, heightPt, afterChars, positionPt?, float? }`。
`file` 相对 spec 所在目录（标定用的 8×8 纯黑 PNG 由 `scripts/make-png.ts` 生成 ——
图片是二进制，扔进仓库就没人说得清它有几个像素，而「像素数与显示尺寸无关」正是要证明的事）；
`afterChars` 是「插在第几个字之后」；`float` 给 `{ relativeH, relativeV, leftPt, topPt, behindDoc? }`
就转成 `wrap="none"` 的浮动图。**同一段里的多张图是倒序插的**（每插一张就多一个字符，
正序会把后面的偏移顶歪），`RelativeHorizontalPosition` **必须在 `Left` 之前设**
（`Left` 是相对当前参照框量的，顺序反了 Word 会按新框重新解释已经写好的坐标）。

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
- `pages[].images[]`：**图片的落点**（外接矩形，`yBottom` 是底边）。它不走
  `getTextContent()`（那一路只吐 show-text 的产物），而是照着算子表把 `q` / `Q` / `cm`
  演一遍 CTM 读出来的 —— PDF 里图片没有自己的坐标，位置与大小全在矩阵里。
  **没有图片的页整个字段不出现**，所以十几份纯文字真值重抽一遍仍逐字节相同
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
