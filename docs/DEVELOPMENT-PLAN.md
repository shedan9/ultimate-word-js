# ultimate-word-js 开发计划

> 目标：自研布局引擎的 Word（OOXML）在线预览 / 编辑库。保真度由自己的引擎决定，不依赖浏览器排版；同时提供强交互 API（定位、装饰、自定义元素锚定、数据绑定）。
>
> 定位：**中文公文 / 周报 / 总结 / 报告类文档的高保真引擎**，不是通用 Word 引擎。范围收敛是这个项目可行的唯一前提。

---

## 0. 技术决策（先定死，避免中途返工）

### 0.1 语言：TypeScript 为主，WASM 作为叶子依赖

| 选项 | 结论 | 理由 |
|---|---|---|
| 全 TS | ✅ 采用 | 排版热点是宽度测量，CJK 宽度缓存命中率极高；50 页文档全量排版 <10ms 量级。渲染瓶颈在 DOM 写入，不在计算。 |
| 全 Rust → WASM | ❌ 不采用 | WASM 操作 DOM 必须跨边界回 JS，负收益；单人业余项目迭代速度是第一约束。 |
| TS + WASM 叶子模块 | ✅ 采用 | 字形整形（harfbuzzjs）、字体表解析用现成 WASM/JS 包，不自己写。 |

**回头触发条件（写死，不凭感觉切换）**：单文档 > 800 页，且增量排版实现后 P95 重排仍 > 100ms → 将 `@uw/layout` 的 `linebreak` + `measure` 抽成 Rust→WASM。为此这两个模块的接口**从第一天就设计成纯数据进出**（`Uint32Array` 码点 + `Float64Array` 宽度 → `Int32Array` 断点），不持有任何 JS 对象引用，保证可平移。

### 0.2 单位系统

内部**统一用 twips（1/1440 英寸）**做浮点运算，因为 OOXML 原生就是 twips/half-point/EMU。

- 禁止在布局过程中出现 px；px 只在渲染层最后一次转换产生
- 转换常量集中在 `@uw/core/units.ts`：`EMU_PER_TWIP = 635`、`TWIP_PER_PT = 20`、`TWIP_PER_INCH = 1440`
- 累加误差：行高、页高累加全程用 twips 整数或定点数，避免 px 浮点漂移导致最后一行溢出

### 0.3 渲染后端

**DOM 为主，Canvas 为可选后端**（打印/导出图片/超长文档降级）。两者共享同一份布局结果（`LayoutResult`），渲染器只是把坐标画出来。

DOM 渲染的关键取舍：

- **不要一字一 span**（DOM 爆炸）。粒度 = "一行内的一个 run 片段" → 一个绝对定位元素
- 需要逐字微调 x（两端对齐、标点挤压、中西文间距）时，用 **SVG `<text x="x1 x2 x3 ...">`** —— SVG text 原生支持逐字形 x 数组，一个元素搞定精确定位，这是 HTML 做不到的
- 页面虚拟化：视口 ±2 页实际渲染，其余用占位盒 + `content-visibility: auto`

---

## 1. 架构与包划分

pnpm workspace monorepo：

```
packages/
  core/          @uw/core        单位、几何、错误、事件、日志、类型
  ooxml/         @uw/ooxml       OPC(zip) + XML → 原始 OOXML 树
  model/         @uw/model       文档模型 + 样式级联 + 编号 + 变更事务
  fonts/         @uw/fonts       字体表解析、度量、替换表、缓存
  layout/        @uw/layout      ★ 布局引擎：测量/断行/段落/表格/分页/域
  render-dom/    @uw/render-dom  绝对定位 DOM 渲染器
  render-canvas/ @uw/render-canvas
  view/          @uw/view        视口、滚动、虚拟化、命中测试、装饰层、Overlay
  editor/        @uw/editor      选区、光标、IME 输入、命令、undo/redo
  serialize/     @uw/serialize   模型 → docx 回写
  ultimate-word/ ultimate-word   门面包，重导出 + 默认组装
  react/         @uw/react       React 组件封装
apps/
  playground/                    Vite 8 调试台（可视化布局盒、逐帧排版）
  fidelity/                      保真度对比工具（见 §6）
```

**依赖方向严格单向**：`core ← ooxml ← model ← layout ← render-* ← view ← editor`。
`layout` **不得** import 任何 DOM API（除 `OffscreenCanvas` 通过注入的度量接口），保证可在 Worker / Node 里跑，也是未来换 Rust 的前提。

### 1.1 工程栈（2026 主流）

| 用途 | 选型 |
|---|---|
| 语言 | TypeScript 7.x（Go 原生编译器），`strict` + `exactOptionalPropertyTypes` + `noUncheckedIndexedAccess` + `erasableSyntaxOnly` |
| 运行时 | Node 24 LTS（经 fnm 管理）。`.ts` 直接用 Node 原生类型剥离跑，工具脚本不需要 tsx |
| 包管理 | pnpm 11 workspace（corepack 托管，`packageManager` 字段锁版本） |
| 构建 | tsdown（rolldown）产库；Vite 8 跑 playground |
| 任务编排 | Turborepo（缓存 build/test） |
| 测试 | Vitest（单测）+ Vitest Browser Mode / Playwright（真实字体度量必须在真浏览器跑） |
| Lint/Format | Biome（单工具替代 ESLint+Prettier） |
| 版本发布 | Changesets |
| 文档 | VitePress + Typedoc |
| zip | fflate |
| XML | fast-xml-parser（v0）；若 profiling 显示瓶颈再换手写 pull parser |
| 断行 | css-line-break（UAX#14）+ 自研中文禁则层 |
| 字体 | fontkit（表解析 + 基础整形）。**不引 harfbuzzjs** —— 其价值在阿拉伯/印度语系复杂整形，属本项目非目标 |

---

## 2. 核心技术难点与解法（这部分决定成败）

### 2.1 字体度量 —— 最容易被低估的一项

Word 的单倍行距**不是** CSS `line-height: normal`。Word 用字体 `OS/2` 表的 `usWinAscent` / `usWinDescent`（部分场景用 `hhea` 的 ascender/descender/lineGap），乘以字号得到行高。浏览器各引擎对 normal 的算法不一致。

**结论：字体表解析不是可选项，是 Phase 1 的必做项。**

#### ✅ Phase 0 穿刺已测定的行高公式（2026-08-13）

实测样本见 `apps/fidelity/fixtures/spike-lineheight-01|02`，验证脚本 `../apps/fidelity/src/spike-lineheight.ts`，
实现在 `@uw/fonts` 的 `lineMetrics()`。**按脚本分两条路**：

| 行的内容 | 单倍行距行高 |
|---|---|
| 含东亚文字 | `(usWinAscent + usWinDescent) × 1.3 × 字号 / unitsPerEm`，**不加**外部行距 |
| 纯拉丁文字 | `(usWinAscent + usWinDescent + GDI 外部行距) × 字号 / unitsPerEm` |

其中 GDI 外部行距（TEXTMETRIC.tmExternalLeading）= `max(0, hhea.lineGap - (win 跨度 - hhea 跨度))`。

13 个样本最大误差 **0.132 pt**（含 ~0.1pt 的 PDF 坐标取整），远低于 1pt 判据。

两个坑：

- **那个 1.3 是乘在字体度量上，不是「行高 = 1.3 × 字号」。** 宋体家族的 `unitsPerEm` 是 256、
  win 跨度恰好 1.0 em，两种假设在它们身上完全重合；要用微软雅黑（1.3198 em → 实测 1.71 em）
  与等线（1.0420 em → 实测 1.35 em）才能分开
- **中文版 Word 的 Normal 模板默认开着行网格**（39 行 / linePitch 312 twips = 15.6pt），
  基线会被吸到 15.6pt 的整数倍上，把字体度量的差异整个盖掉。做度量实验必须显式关掉网格
  （`PageSetup.LayoutMode = wdLayoutModeDefault`），否则量到的全是网格间距

**未决**：这 30% 的额外行距在基线上下如何分配（决定行内基线的确切位置）。
Phase 2 之前要再做一次「首行基线到版心顶」的穿刺定下来。

#### 度量三级策略

| 级别 | 来源 | 精度 | 说明 |
|---|---|---|---|
| ① 真实字体文件 | fontkit 解析 `OS/2`/`hhea`/`hmtx`/`cmap`/`GPOS` | 与 Word 一致 | 主力路径 |
| ② 度量包（metrics pack） | 随库分发的纯度量 JSON | 与 Word 一致 | 见下，解决中文字体缺失 |
| ③ `canvas.measureText` | 浏览器字体引擎 | 近似 | 兜底，仅未知字体 |

字体文件来源：① docx 内嵌字体（`w:embedRegular`，注意需按 GUID 做 obfuscation 解混淆）② 用户注册的 webfont（我们自己 fetch 字节）③ 系统字体 —— 浏览器**不提供**字节，只能走级别 ②/③。

#### 度量包：中文字体缺失的实际解法

#### 支持字体清单（首批）

**A 类 · 中文（度量包 + 替代字体渲染）** —— 无开源度量兼容替代，只能抽度量

| 中文名 | 英文名 | 文件 | 渲染替代 |
|---|---|---|---|
| 宋体 | SimSun | `simsun.ttc` | Noto Serif CJK SC |
| 仿宋 | FangSong | `simfang.ttf` | Noto Serif CJK SC |
| 黑体 | SimHei | `simhei.ttf` | Noto Sans CJK SC |
| 楷体 | KaiTi | `simkai.ttf` | Noto Serif CJK SC / LXGW WenKai |

暂不支持：`仿宋_GB2312`、`楷体_GB2312`、`方正小标宋简体`。
（备注：GB/T 9704 公文格式里正文常指定 `仿宋_GB2312`，与 `仿宋` 是两款不同字体、度量不同。若后续真实公文大量使用，按同样流程加度量包即可。）

**B 类 · 拉丁核心（直接打包开源度量兼容克隆）** —— 度量与渲染一并解决，且无授权问题

| Word 字体 | 度量兼容克隆 | 许可 |
|---|---|---|
| Times New Roman | Liberation Serif | OFL |
| Arial | Liberation Sans | OFL |
| Courier New | Liberation Mono | OFL |
| Calibri | Carlito | OFL |
| Cambria | Caladea | OFL |

> 「度量兼容」= 逐字宽度与原字体相同，专为替换设计。这比 A 类的度量包更优：跨平台不仅分页一致，字形外观也基本一致。

**C 类 · 拉丁次要（仅度量包，渲染回退系统 serif/sans）**
Georgia、Verdana、Tahoma、Segoe UI、Aptos。使用频率低，不值得增加包体积。
（Aptos = Microsoft 365 自 2024 起的默认主题字体；当前开发机为永久授权版 Word，默认仍是 Calibri/Cambria，暂无此字体文件。）

**D 类 · 符号字体（必做）**
`Symbol`、`Wingdings` —— 项目符号的实际载体：`numbering.xml` 里实心圆点是 Symbol 的 `0xB7`，实心方块是 Wingdings。不支持则所有列表 bullet 渲染错误。

> ⚠️ 坑：二者是 **symbol-encoded** 字体，使用 `(3,0)` cmap 子表，字符映射到 **U+F020–U+F0FF 私用区**。docx 中写 `w:char="F0B7"`，需减去 `0xF000` 或直接查 (3,0) 表 —— 用常规 `(3,1)` Unicode cmap 查会全部落空。

#### 度量包机制

A 类与 C 类字体在 Linux/Mac 上缺失。关键认识：**我们需要的只是度量，不是字形。**

离线从真实 Windows 字体抽取纯度量包随库分发：

```jsonc
{ "name": "仿宋_GB2312", "unitsPerEm": 1000,
  "os2": { "winAscent": 880, "winDescent": 120 },
  "defaultAdvance": 1000,                    // CJK 绝大多数全角等宽
  "exceptions": { "0020-007E": [...] } }     // 只有 ASCII 等比例段需逐字列
```

因为 CJK 字体里汉字几乎全是 1em 等宽，例外只有 ASCII 一小段，**单个字体的度量包约 1–2 KB**。效果：非 Windows 平台用替代字体**渲染**、用真实度量**排版** → 断行点与页数和 Word 完全一致，仅字形外观不同。这比想办法凑齐字体授权现实得多。

配套仍需「字体名 → 渲染用替代字体」映射表（见 A 类表格的"渲染替代"列）。

#### `w:rFonts` 按脚本分桶 —— 中文文档度量出错的头号原因

`w:rFonts` **不是**给整个 run 指定一款字体，而是同时挂四个属性：

```xml
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"
          w:eastAsia="宋体" w:cs="Times New Roman" w:hint="eastAsia"/>
```

引擎必须**逐字符**判断其所属脚本桶（ascii / hAnsi / eastAsia / cs），再选对应字体取度量。这正是"汉字用宋体、数字与英文用 Times New Roman"的实现机制。

- 分桶依据是字符的 Unicode 区段；歧义区段（如全角标点、部分符号）由 `w:hint` 决断
- 因此**一个 run 内可能横跨多款字体**，`ShapedRun` 的最小单位是「同字体的连续字符段」，不是 run
- 主题字体：`w:asciiTheme="minorHAnsi"` 等需经 `theme1.xml` 的 `<a:latin>` / `<a:ea>` 解析后才拿到真实字体名，务必在样式级联阶段就解析掉

#### 关于 pretext（github.com/chenglou/pretext）—— 评估结论：不作为核心

它的设计目标是「不触发 reflow 地预测**浏览器**会怎么排」，真值取自 canvas measureText；我们的目标是预测 **Word** 怎么排，必须主动偏离浏览器。硬冲突项：

- 不解析任何字体表 → 拿不到 `hmtx`/`OS/2`
- `lineHeight` 是**调用方传入的参数**（原文 "lineHeight stays a layout-time input"）→ 而行高恰是我们最难算的**输出**
- 明确不提供逐字形 x（"not exact glyph-position data for ... x-coordinate reconstruction"）→ 两端对齐 / 标点挤压 / 中西文间距无法实现
- 制表位仅 `tab-size: 8`；断行是 CSS 的 `word-break: keep-all` 语义而非 UAX#14 + 禁则

**可借鉴处**：级别 ③ 兜底层。它在 grapheme 边界、代理对、emoji ZWJ、`Intl.Segmenter` 分段与缓存上做得扎实，这层可直接依赖或抄架构。

### 2.2 中文排版（差异化核心，docx-preview / OnlyOffice 都没做全）

必须实现的五件事：

1. **行网格 `w:docGrid`**
   - `type="lines"` / `linesAndChars` / `snapToChars`
   - 公文的「每页 22 行 × 每行 28 字」就靠这个。行基线要吸附到网格，段落行高 = `linePitch` 的整数倍
   - 这是公文保真的第一名，优先级高于表格

2. **避头尾（禁则）**
   - Word 有内建首禁则集（`,。、；：？！》」』）】…—` 等）和尾禁则集（`《「『（【` 等）
   - 处理策略：**先尝试压缩（挤压标点），压不下再回退（把前一字符推到下一行）**，Word 的实际策略是压缩优先
   - 字符集可被 `w:settings` 里的 `w:noLineBreaksAfter` / `w:noLineBreaksBefore` 覆盖

3. **标点挤压 / 溢出**
   - 全角标点在行首行尾的半角化压缩（Word 内部按 0.5em 压）
   - `w:overflowPunct`（允许标点溢出边界）默认开启 —— 行尾句号可以吐出版心

4. **中西文自动间距 `w:autoSpaceDE` / `w:autoSpaceDN`**
   - CJK 与拉丁字母 / 数字之间自动插入 1/8 em 间隙，默认开启
   - 这个不做，中英混排的行长就永远对不上

5. **首行缩进的字符单位 `w:firstLineChars`**
   - 值是 1/100 字符，「首行缩进 2 字符」= `firstLineChars="200"`，实际宽度 = 2 × 当前字号的全角宽，**不是** 固定 twips

### 2.3 分页

- 孤行寡行控制 `w:widowControl`（默认开）
- `w:keepNext` / `w:keepLines` / `w:pageBreakBefore`
- 表格跨页拆行、`w:cantSplit`、表头行 `w:tblHeader` 重复
- 脚注：**不动点问题** —— 加脚注挤走正文，正文变了脚注归属页也变。解法同 §2.4 的迭代收敛

### 2.4 域（PAGE / NUMPAGES / TOC）的循环依赖

页码依赖布局 → 目录长度依赖页码 → 目录变长又改变布局。

**解法：迭代到收敛。**

```
layout()                        // pass 0：域取上次结果或占位值
for i in 1..MAX_ITER (=5):
    resolveFields()             // 用 pass i-1 的布局结果算 PAGE/TOC
    layout()
    if pageCount 与 field 文本均未变化: break
    if 检测到 A→B→A 振荡: 取页数较大者冻结，退出
```

- `MAX_ITER = 5`（Word 实际 2–3 次收敛）
- 必须有**振荡检测**，否则遇到临界文档会死循环
- 目录项文本长度变化要走「同一次 pass 内只允许增长」的阻尼策略

### 2.5 增量排版

三级脏标记，否则编辑态每敲一个字全量重排：

| 层级 | 触发 | 重算范围 |
|---|---|---|
| 行级 | 段内文本修改 | 从受影响行到段落末，若段落总行数不变则**不向上冒泡** |
| 段级 | 段落属性变化、行数变化 | 该段 + 后续分页 |
| 节级 | 页面设置、分节符变化 | 该节全部 |

段落布局结果缓存 key = `hash(段落内容 + 解析后属性 + 可用宽度)`。

### 2.6 编辑与 IME（中文输入的生死线）

- **不用 contenteditable 承载正文**。正文是绝对定位的只读 DOM
- 用一个 1×1 px 的隐藏 `contenteditable` 或 `textarea` 跟随光标，接管 `beforeinput` / `compositionstart|update|end`（Monaco / CodeMirror 6 的做法）
- 组合期（拼音输入中）在光标处渲染一个**临时预览层**，不进模型；`compositionend` 才提交事务
- 命中测试完全自研：布局层已有每个字形的精确 x/y/w/h，二分查找即可，不依赖 `caretPositionFromPoint`
- 选区：编辑态用自研 overlay 矩形；只读态额外提供「原生可选文本层」模式，让 Ctrl+F、划词、屏幕阅读器可用

---

## 3. 公开 API 设计（交互能力是本库的卖点）

```ts
import { UltimateWord } from 'ultimate-word';

// ── 加载 ────────────────────────────────────────────
const doc = await UltimateWord.load(arrayBuffer, {
  fonts: { substitutes: { '仿宋_GB2312': 'FangSong' } },
});

const view = doc.mount('#container', {
  mode: 'preview',          // 'preview' | 'edit'
  zoom: 'fit-width',
  pageGap: 16,
});

// ── 查询与定位 ──────────────────────────────────────
doc.query('paragraph[styleId=Heading1]');        // 类 CSS 选择器
doc.find('签发人');                                // → DocRange[]
doc.locate({ page: 3, x: 1200, y: 4300 });        // 坐标 → 模型位置
view.rectsOf(range);                              // 模型位置 → 屏幕矩形
view.scrollTo(range, { align: 'center' });

// ── 装饰与自定义元素 ────────────────────────────────
const d = view.decorate(range, { className: 'hl', style: { background: '#ff0' } });
d.dispose();

view.overlay(anchor, element, {                   // 锚定任意 DOM
  placement: 'right-of-line',
  follow: true,                                   // 重排后自动跟随
});

// ── 数据前端控制（模板填充）──────────────────────────
doc.bindings.set('applicant', '张三');             // 映射到 w:sdt 内容控件
doc.bindings.apply();                             // 触发增量重排

// ── 编辑命令 ────────────────────────────────────────
doc.tx(t => {                                     // 一个事务 = 一个 undo 单元
  t.insertText(pos, '正文内容');
  t.setParagraphProps(pos, { firstLineChars: 200 });
});
doc.undo(); doc.redo();

// ── 事件 ────────────────────────────────────────────
doc.on('layout:done', ({ pageCount, duration }) => {});
view.on('selection:change', sel => {});
view.on('click:element', e => {});

// ── 导出 ────────────────────────────────────────────
await doc.toDocx();      // → Blob
await view.toPNG(3);     // 第 3 页
```

设计原则：

- **模型位置（`DocPosition`）与屏幕坐标严格分离**，两者互转由 view 层提供，模型层不知道像素
- 所有装饰 / overlay 返回 `Disposable`，不泄漏
- 数据绑定优先走 OOXML 原生的**内容控件 `w:sdt`**，回写 docx 时天然兼容 Word

---

## 4. 分阶段路线图

每个阶段都以**可演示的产物**结束。业余时间推进，不设日期，只设完成判据（DoD）。

### Phase 0 — 地基与验证性穿刺 ✅
- ~~pnpm + Turborepo + tsdown + Vitest + Biome 骨架~~ ✅；~~CI~~ ✅ `.github/workflows/ci.yml`
- ~~**穿刺实验**：手写一个 5 段中文文档，用自己算的字体度量排一页，与 Word 导出 PDF 的实际坐标逐行比对~~ ✅
- **DoD**：~~证明「读 OS/2 表算行高」能把单页行基线误差压到 < 1pt~~ ✅ 实测最大误差 **0.132 pt**（见 §2.1）

### Phase 1 — OOXML 解析 + 文档模型 + 样式级联 🚧
- ~~OPC 容器（fflate）、关系解析、part 索引~~ ✅ `@uw/ooxml`：解包 + 内容类型 + 关系 +
  保序 XML 纯数据树 + 反向序列化。`gongwen-01.docx` 全部 11 个部件语义 round-trip 通过
- ~~`styles.xml` / `theme1.xml`~~ ✅ · `document.xml` 正文节点树 / `numbering.xml` / `settings.xml` / `fontTable.xml` 待做
- ~~样式级联：`docDefaults → styles.xml(basedOn 链，含循环检测) → 直接格式`~~ ✅ `@uw/model/cascade.ts`
  - 两个**写明的洞**：编号那一层（`numbering.xml` 每级自带 pPr/rPr）留到 Phase 5；
    toggle 属性（b / i / caps…）在样式层之间的 XOR 语义（§17.7.3）按「后者覆盖」处理 ——
    没有 Word 真值样本能验证 XOR 的边界，照规范硬写一个测不了的实现比留个洞更危险
  - 表格条件格式那一层的位置也还空着，Phase 4 再接
- ~~属性解析成**扁平化的 `ResolvedRunProps` / `ResolvedParaProps`**，布局层不再碰 XML~~ ✅
  单位在解析处就转 twips；两个**故意不转**的：`w:line`（刻度取决于 `w:lineRule`）与
  `w:*Chars`（1/100 字符，字号级联后才知道实际宽度）
- **DoD**：任意公文 docx 能 dump 出完整的解析后属性树，与 Word「显示格式」面板抽查一致

### Phase 2 — 度量 + 段落布局 + 单页 DOM 渲染
- `@uw/fonts`：fontkit 解析、度量缓存（两级：字体级 Map + 全局 LRU）、替换表
- 断行：UAX#14 基础 + 中文禁则 + 标点挤压 + 中西文间距
- 段落布局：对齐（含分散对齐）、缩进（含字符单位）、行距（单倍/固定值/最小值/多倍）、段前段后、制表位（左/中/右/小数点/前导符）
- 行网格 `docGrid` 吸附
- DOM 渲染器 v1：绝对定位 + SVG text 逐字 x
- **DoD**：单页纯文本公文，与 Word 截图叠加对比，字形位置误差 < 1pt

### Phase 3 — 分页 + 分节 + 页眉页脚
- 分节符（next page / continuous / even / odd）、每节独立页面设置
- 页眉页脚：首页不同、奇偶不同、`linkToPrevious`
- 分页规则：widow/orphan、keepNext、keepLines、pageBreakBefore
- PAGE / NUMPAGES 域 + 收敛循环（§2.4）
- 页面虚拟化
- **DoD**：20 页真实公文，总页数与 Word 一致，每页首末字一致

### Phase 4 — 表格
- 固定布局（`tblLayout=fixed`）优先，`autofit` 做简化版
- 合并：`gridSpan`（横向）、`vMerge`（纵向）
- 边框冲突解析（Word 的边框优先级规则）、单元格边距、垂直对齐
- 跨页：拆行、`cantSplit`、`tblHeader` 重复表头
- 嵌套表格
- **DoD**：公文常见的「发文单位 / 签发人」表头表格、以及三线表，能正确跨页

### Phase 5 — 列表编号 + 域 + 图片
- `numbering.xml`：多级列表、`numFmt`（含 `chineseCounting` / `chineseCountingThousand`）、`lvlText`、重启规则
- 域：TOC（含收敛）、SEQ、STYLEREF、DATE、HYPERLINK
- 图片：inline 为主，anchor 只做「上下型环绕」，其余环绕类型退化为 inline（明确写进非目标）
- **DoD**：带自动目录的报告，目录页码正确且跳转可用

### Phase 6 — 交互 API（只读态）
- 命中测试、坐标 ↔ 模型位置双向映射
- `find` / `query` / `decorate` / `overlay` / `scrollTo`
- 原生可选文本层模式（Ctrl+F、复制、无障碍）
- 打印（`@media print` + 分页 CSS，或直接走 canvas → PDF）
- **DoD**：能在文档任意段落右侧挂一个 React 批注气泡，滚动 / 缩放 / 重排后位置不飘

### Phase 7 — 编辑态
- 选区模型、光标渲染与闪烁、Shift/Ctrl 移动、双击选词（用 `Intl.Segmenter`）
- 隐藏输入框 + IME 组合层（§2.6）
- 事务系统 + undo/redo（命令合并：连续输入合并为一个 undo 单元）
- 基础命令集：输入/删除/粘贴（纯文本 + HTML + docx 片段）、加粗斜体、段落属性、列表
- 增量排版接入
- **DoD**：能连续用拼音输入法打完一整段中文，光标位置与撤销行为符合直觉

### Phase 8 — 回写 docx
- 模型 → OOXML 序列化，**保留未识别的原始 XML**（round-trip 安全：解析时把不认识的元素原样挂在节点上，回写时吐回去）
- **DoD**：加载 → 不编辑 → 导出，得到的 docx 用 Word 打开无修复提示，且与原文件语义等价

### Phase 9 — 长尾与打磨
- 保真度回归库扩容
- 性能：Worker 化布局、字体预热、首屏时间
- 文档站 + 示例

---

## 5. 明确的非目标（写下来，防止自己失控）

- 浮动对象的**紧密型 / 穿越型环绕**（`wrapTight` / `wrapThrough`，多边形绕排）
- 多栏排版（`w:cols` 多栏）—— Phase 9 之后再议
- 修订痕迹（track changes）的编辑，只做**显示**
- 数学公式 OMML 的排版 —— 转 MathML 交给浏览器
- VML / 旧版图形、SmartArt、图表（`c:chart`）—— 降级为占位图
- .doc（二进制）格式
- 完整 autofit 表格算法
- RTL / 复杂文字（阿拉伯、天城文）

---

## 6. 保真度验证体系（最高杠杆的基础设施，Phase 0 就要建）

**开发机为 Windows 且已安装 Word（已确认）—— 这意味着可以拿到数值级的排版真值，而不是靠肉眼比对。**

```powershell
# apps/fidelity/scripts/export-truth.ps1
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($InputPath)
$doc.ExportAsFixedFormat($OutputPdf, 17)   # 17 = wdExportFormatPDF
$doc.Close($false); $word.Quit()
```

然后用 pdf.js 从 PDF 里抽取**每个文本片段的 transform 矩阵**，得到 Word 排出来的精确 (x, y, fontSize, text)。

这样得到的不是像素图，是**坐标真值**，可以直接写断言：

```ts
expect(layout).toMatchWordTruth('fixtures/gongwen-01.docx', {
  tolerance: { x: 0.5, y: 0.5, unit: 'pt' },   // 0.5pt 容差
  assertPageCount: true,
  assertLineBreaks: true,                       // 每行首末字符必须一致
});
```

分级断言（从松到严，逐阶段收紧）：

| 级别 | 断言内容 | 目标阶段 |
|---|---|---|
| L0 | 总页数一致 | Phase 3 |
| L1 | 每页首行 / 末行文本一致 | Phase 3 |
| L2 | 每行的断行点一致（行首末字符） | Phase 2 |
| L3 | 每行基线 y 误差 < 0.5pt | Phase 2 |
| L4 | 每个 run 片段起始 x 误差 < 0.5pt | Phase 4 |

**语料库**：从现在开始攒 `fixtures/`，按类型分组（红头文件 / 请示 / 通知 / 会议纪要 / 周报 / 带目录的报告 / 复杂表格）。每修一个 bug，就把那个文档加进去。这个库的价值最后会超过代码本身。

CI 上无 Word，所以真值 PDF 与抽取结果**提交进仓库**（`fixtures/*.truth.json`），本地重新生成。

---

## 7. 立即可执行的下一步

1. ~~工具链~~ ✅ fnm 1.39 + Node 24.19.0 + npm 11.17 + pnpm 11.21（corepack）。
   fnm 初始化已追加到 `$PROFILE`；`--use-on-cd` 会读项目里的 `.node-version`
2. ~~`git init` + `../.gitignore` + `.node-version` + `packageManager`~~ ✅
3. ~~monorepo 骨架：`core` / `ooxml` / `model` / `fonts` / `layout` / `render-dom` + `../apps/playground`~~ ✅
   （`core/units.ts` 与 `fonts/metrics.ts` 已是真实实现，其余是占位）
4. ~~`../apps/fidelity`：PowerShell 导出 PDF → pdf.js 抽坐标 → JSON 真值~~ ✅ 见 `../apps/fidelity/README.md`
5. ~~Phase 0 穿刺：验证行高算法误差 < 1pt~~ ✅ 0.132pt，结论写进 §2.1

接下来（按顺序）：

6. **补基线位置穿刺** ⏸ **卡在 Windows** —— 定下东亚行高里那 30% 额外行距在基线上下如何分配（§2.1 未决项）。
   做法：量「首行基线到版心顶」的距离，与 ascent 预测对比。这个不定，Phase 2 的行盒就摆不准。
   要 Word COM + `C:/Windows/Fonts`，Mac / Linux 上会被 `platform.ts` 以退出码 2 拦下
7. ~~**CI**（GitHub Actions）：`typecheck` + `test` + `biome check`~~ ✅ 真值不在 CI 生成，只读仓库里的 `*.truth.json`
8. **Phase 1 开工** 🚧：~~`@uw/ooxml` 的 OPC 容器 + part 索引~~ ✅ → 下一步 `@uw/model` 的样式级联。
   起手就用 `fixtures/gongwen-01.docx`（已入库）当解析目标
9. 语料库扩容：把真实公文丢进 `../apps/fidelity/fixtures`，每修一个 bug 加一个文档

第 6 步优先于任何**布局**代码 —— 与第 4 步同理，没测准的东西不要拿来当地基。
但它不挡 Phase 1：解析与样式级联跟基线位置无关，可以在非 Windows 机器上一直做到 Phase 2 门口。
