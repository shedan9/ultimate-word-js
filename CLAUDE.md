# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 这个项目是什么

自研布局引擎的 Word（OOXML）在线预览 / 编辑库，定位是**中文公文 / 周报 / 报告类文档的高保真引擎**。
保真度由自己算的排版决定，**不依赖浏览器排版** —— 所以「用 CSS 让它看起来差不多」永远不是正确答案。

当前进度：Phase 0 已完成（地基 + 行高穿刺 + 基线穿刺 + CI），Phase 1 的**解析链已完整**，
**Phase 2 全部做完了**（DOM 渲染器 v1 已落地），**Phase 3 与 Phase 4 也全部做完了** —— `layoutDocument()`
把段落与表格摞进一页页的版心，每一行都有了页号与 y，`LineLayout` 的 `baseline` 终于能拼成
绝对坐标；**页眉页脚**（选份 + 定位 + 反过来挤版心）三条几何规则已用真值标定。
**行高走东亚还是拉丁规则**也标定完了 —— 看的是**字体**不是字符（见下），
这条修掉了一个「每一行都差 30%」的老错。**宽度那一维**（歧义字符与空格的分桶、
中西文自动间距）也标定完了，同样修掉一条老错：自动间距是 **1/4 em 不是 1/8**（见下）。
**卡口没有了**：横向（断行）与纵向（基线 y）现在都能与真值逐行比。
Phase 5 的**列表编号**已经从 `numbering.xml` 一路通到首行几何，
同阶段的**域**结构还原（界桩配对 + 指令解析 + HYPERLINK）与**求值**（PAGE / NUMPAGES /
SECTIONPAGES 迭代到自洽）都做完了，TOC / SEQ 的求值还没写。
**图片**也通了（解析 → 收字节 → 占位 → 画，四层各一段，见下），**几何也用真值标定完了**。Phase 4 的**表格**：属性 + 级联（含 `w:tblStylePr` 条件格式）在 model 层，
列宽 + 每格的 x 与可用宽 + 格内段落 + **边框冲突解析**在 layout 层，跨页按**行**拆，
一行放不下时还会从**行间**切开（**拆行**，见下）。**表格这一层全部标定完了** ——
几何、条件格式、格线冲突、拆行四问都有了真值。造样本的工具从此会造表
（`make-fixture.ps1` 的 `kind: "table"`），真值也从此读得到**画出来的线**
（`truth.json` 的 `pages[].rules[]`），见下。
**Phase 6 的地基也起了**：`buildLayoutIndex()` 把「点 ↔ 模型位置」的两个方向做完了（见下），
剩下的是屏幕坐标那一跳（`@uw/view` 还没建）。
真实实现：`@uw/core`（单位 / 错误 / 诊断）、`@uw/ooxml`（OPC 容器 + XML 树）、
`@uw/model`（样式级联 + 主题字体 + 正文节点树 + 分节 + 设置 + 字体表 + 制表位 + **模型位置（`DocPosition`）** + **编号（解析 + 计数器 + 编号文字 + 接进级联）** + **表格（属性 + 级联 + 条件格式）** + **域（界桩配对 + 指令解析 + HYPERLINK）** + **页眉页脚部件** + **图片（外框 + blip 引用 + 裁剪 / 旋转 + 浮动锚点 + 字节表）**）、
`@uw/fonts`（行高规则 + 脚本分桶（**歧义字符 / 中性字符两条都实测**）+ 度量包 + 注册表 + `TextMeasurer`）、
`@uw/layout`（item 流 + 断行 + 缩进 / 对齐 / 制表位 / 列表编号 + **中西文自动间距** + 行高与网格吸附 + **行内基线** +
**表格列宽与格内几何 + 边框冲突解析** + **分页（含表格拆行）** + **域求值** + **页眉页脚** + **对象占位与浮动定位** + **布局索引（命中测试 / range → 矩形 / 光标）**）、
`@uw/render-dom`（**元素树 → SVG / DOM**，见下）。

**渲染器**（`packages/render-dom`）是流水线的出口，**px 只在这一步出现**。
一页一个 `<svg>`，**viewBox 的单位是 pt** —— 与 `fixtures/*.truth.json` 同一套坐标
（原点纸左上角、y 向下），属性里读到的 `y="119.05"` 就是真值里那个数，比对不用换算。
逐字 x 走 `<text x="x1 x2 …">`，粒度 = 一行里的一个 run 片段（不是一字一元素）。
**缩放只改 `<svg>` 的 width / height，viewBox 一个字不动** —— 架构 §4.1 的
「缩放永不触发重排」就落在这一行属性上。两个入口：主入口只到「纯数据元素树 → 标记文本」，
**一个 DOM API 都不碰**（单测跑纯 Node，`preview` 拿它落盘成 HTML，将来 Worker 里也走这条）；
`@uw/render-dom/dom` 才建真 DOM，`Document` 是注入的。分开不是洁癖：workspace 包的
`exports` 直接指向 `src/*.ts`，主入口牵进 `dom.ts` 会逼着 `@uw/fidelity` 打开 `lib: ["DOM"]`。
表格分三遍画（底纹 → 文字 → 线），**因为格线是共享的** —— 逐格画完再画下一格，
后一格的底纹会盖掉先画的半条线；共享的线还按几何位置去重，只画一次。
渲染层**不回头查模型**：要画的视觉属性由 `LineFragment.style` 从布局带过来
（包依赖单向，且 Worker 化之后主线程手上只有 `DocumentLayout` 这一份数据）。
页眉页脚是与版心 `<g>` **平级**的另外两个 `<g>`（坐标相对纸左上角）—— 它们不在版心里，
版心正是被它们挤出来的。浮动对象（印章 / 水印）同样与版心平级，**衬于文字下方的画在正文之前、
浮于上方的画在最后** —— SVG 里的「层」就是画的先后。
未画：run 级高亮（model 没解析）、可选文本层、增量更新。
画法里没有真值的常数（下划线 / 删除线的位置粗细、上下标升降量、前导符点距）
关在 `packages/render-dom/src/uncalibrated.ts` ——
**它们一个都不改坐标**，所以 L2/L3/L4 全绿也证明不了它们对。

**布局索引**（`packages/layout/src/layout-index.ts`，Phase 6 的地基）：`buildLayoutIndex(doc)`
把整份 `DocumentLayout` 摊平成一张「行表」，答四件事 —— `positionAt(点)` 命中测试、
`rectsOf(range)` 装饰要画的矩形、`caretRect(位置)` 光标、`compare` 文档序。
架构 §4 的 ①↔② 就是它；②↔③（px）归将来的 `@uw/view`，**它要知道缩放与滚动，那是每个视图各不相同的**。
它**在消费侧现建**（带方法的对象过不了结构化克隆），只吃 `DocumentLayout` 这一份纯数据。
四处容易搞反：① **不是二分查找** —— 表格让同一个 y 上并排坐着好几行，行序与 y 序不一致，
实现是「按页分桶线性扫，先比纵向距离再比横向」；② **`DocPosition` 是三个字段**
（run + 内容片段下标 + 片段内偏移，定义在 `@uw/model` 的 `position.ts`）：run 的内容是一列
**没有 id** 的片段，「run 内的全局偏移」要有模型才算得出来，而索引这一侧没有模型；
③ **编号与域结果不可定位**（源位置是 -1），点在它们身上给的是旁边真实的字，
整段选中时编号不进 range、域结果只在被整个包住时才算；④ **重复表头收进索引但标出来**
（`IndexedLine.repeated`）—— 点它能把光标放进文档里那一份（Word 也是这样），
但可选文本层必须跳过它，否则每页复制出一遍表头。
为它给 `LineFragment` 补了 `contentIndex` / `offset` 并让**片段不跨内容片**：
有了这条保证，「片段里第 k 个码点」的位置 = 首字位置 + 前 k 个码点的长度，
不必给每个字形再存一份位置。摊平那一段的几何与渲染层的 paint.ts **一字不差**，
两边各写一遍是因为依赖方向单向 —— 命中测试不该逼着调用方先渲染一遍。
**没有新的标定**：字摆在哪由已经标定完的那几层决定，判据是单测不是真值残差。

**语料体检**（`apps/fidelity/src/corpus-report.ts`，`pnpm --filter @uw/fidelity corpus`）：
把 `fixtures/` 里非 spike 的文档整份排一遍，与它的 `truth.json` 比 L0（页数）/ 每页行数 /
L2（每行文字），并把诊断按 code 汇总。**它不是测试** —— 新语料进来时页数对不上是常态，
它的用处是「差在哪」一眼看清。它要**自己重做一遍分行**：真值的一「行」是按 y 分桶拼出来的，
不认识段落也不认识表格，所以这边也得把正文 + 表格 + 页眉页脚一起摊平再按同样的规则分桶
（只收正文段落的话表格类文档会得出「我们 234 行 / Word 505 行」这种毫无意义的数字）。
表格类文档的 L2 百分比**不可信**：pdf.js 不吐连续空格，「版    本：」与「版 本：」比不上。

它照出来的两条**制表位**结论（都在 `linebreak.ts` / `items.ts`）：
① **非左对齐制表位的推进量要减去它后面那段**（`alignedTabWidth`，往后看到下一个制表位或段末为止）。
Word 的目录条目是「标题 → 右对齐制表位 → 页码」，那个停靠点**正好等于版心宽** ——
按「推进到停靠点」估宽的话这一行到此已经吃满，页码只能换行，每条目录都排成两行；
② **制表位也有字体**（`TabItem.font`，按 ASCII 桶取）。它不占字形却**参与行高** ——
只有一个制表位的那一行（目录、签发人栏）行高全靠它，缺了就拿空字体名去问度量器、
退到等宽近似，顺带每份文档报一条 `font-missing 字体「」`。

**两个「看得见」的出口**（改完布局或画法值得跑一眼）：
`pnpm --filter @uw/fidelity preview [name] -- --truth --debug` 把 fixture 画成
`apps/fidelity/out/*.html`（真值基线红虚线、我们的蓝实线，重合即对；产物不入库）；
`pnpm --filter @uw/playground dev` 拖一份 docx 进去就画，带缩放与版心 / 行盒开关。
preview 报的「最大差 x pt」**不是保真度指标** —— 它按行序号硬配对，一行断错后面全错位，
真正的判据在 `layout/src/fixture.test.ts` 与各个 spike 脚本里。

**分页**（`packages/layout/src/page.ts`）四处容易搞反：
① `w:sectPr/w:type` 说的是**本节自己**从哪儿开始，不是「下一节怎么开始」（§17.6.22）——
按后者实现整份文档会错开一节；② **keepNext 是「接缝」不是「整块」**：本段末行与下一段首行
同页即可，整段照样能拆，所以实现成「排本段时把接缝高度算进可用高」（`joinHeight`），
按整块原子做会平白把一整段推到下一页；③ **页是惰性开的**（`Flow.page === undefined`），
文末的硬分页符因此不会凭空多出一张空页，`evenPage` / `oddPage` 补的空页才是显式造的
（带 `filler`）；④ `continuous` 只在**页面设置没变**时才真的不换页 —— 一页只能有一个版心框。
每行的绝对基线 = `page.geometry.content.y + PlacedLine.y + LineLayout.baseline`，
gongwen-01 的 18 行与 Word 真值最大差 **0.06pt**（L3 判据 0.5pt），断言在 `fixture.test.ts`。
逐行累加**不需要**「每页重新对齐网格」的修正 —— 网格吸附已经吸在每一行的行高上了。
未做：脚注不占位。

**表格的几何已经标定完**（样本 `spike-table-01/02`，跑 `pnpm --filter @uw/fidelity spike:table`，
落在 `table.ts` 的 `TABLE_RULES`）。这一层此前**一行都没跟 Word 比过** —— 列宽、格内边距、
行高、`w:vAlign`、跨列的可用宽全是照规范推的。一跑就照出一条真错的：
**水平格线占纵向的高、竖格线不占横向的宽**。原来两个方向都按「不占」写，于是每张带框的表
都偏矮：一张 20 行、0.5pt 框线的表少算 10pt，1pt 框线少 20pt，跨页位置一路错下去。
不对称的原因看一眼就明白：**宽度是给定的**（`w:tblGrid` 是 Word 存盘时算完写下的），
边框没地方可占；**高度是算出来的**，边框就能加进去。落到数据上是两个新字段：
`RowLayout.gridAbove`（本行**上边**那条线，已含在 `height` 里）与 `TableLayout.gridBelow`
（表最下面那条，不属于任何一行）。`w:trHeight` 与 `w:vAlign` 量的都是**格线以内**那一段。
顺带验完的几条本来就对：默认单元格边距真的是 108 twips（来自 `Normal Table`）、
`w:tcMar` 覆盖、跨列格按合并后的宽度断行、`w:jc="center"` 与 `w:tblInd`、
`w:vAlign` 的 0 / 一半 / 全部。**剩一处没有模型**：Word 把格内文字再往右挪一点点
（边距 5.4pt 时 0.32pt、20pt 时 0.24pt、0 时 0.59pt），三个数凑不出规则，
关在 `uncalibrated.ts` 的 `TABLE_CELL_TEXT_INSET`（值取 0），也是 `spike:table` 唯一对不上的一项。

**条件格式的层序也标定完了**（同一份 `spike-table-02`）。做法是给每个条件设一个**独一无二的字号**，
于是「这一格最终几号字」= 「层序里最后一个命中它的条件是谁」，从真值的 `size` 直接读得出来 ——
不必从字形宽度反推（那是拿一个未知量去解另一个未知量）。**照出一条反的**：
`CONDITIONAL_ORDER` 里原来写「行带在列带之后」，实测是**列带盖行带**。
注意它与首末那一组**方向相反**（首末行盖首末列），不是一句「行优先」能概括的。
「一格命中哪些条件」是另一回事，用 Word 自己写在 `w:cnfStyle` 上的归属标记验过了，
与实现一致 —— 首行首列被 `tblLook` 排除时不进带、没排除时照样进带、`rowBandSize=2` 按 0 起数分组。

**表格拆行**（`packages/layout/src/table-split.ts`）：一行放不下时从**行间**切开
（格内段落的两行之间、嵌套表格的两行之间），本页一片、下一页接一片；`w:cantSplit`
与**表头行**（每页都要重复一遍，半行表头没有意义）除外。每一格各切各的，
这一片的高度按**最高那一格**算 —— 与不拆行时 `rowHeight()` 取 max 是同一条规则。
切出来的是**两份各自自洽的 `RowLayout`**，不是「一份 + 裁剪窗口」：后者要渲染层加
`clipPath`、要命中测试知道「这一片只露出第几行」、还要一套行内局部坐标。
它修掉的是一个**真会错位**的洞：一行高过整页版心时，原来只能硬塞、内容溢出版心且
后面每页跟着错。

**拆行的四问也标定完了**（样本 `spike-table-04`，跑 `pnpm --filter @uw/fidelity spike:table-split`，
落在 `TABLE_SPLIT_RULES`）。四条原来全是「哪种最省地方」猜的，**三条猜反了**：
① **就地切**（本页剩下多少用多少），不是「整行挪到下一页顶上再切」—— 后者白扔掉本页剩下的
一整块地方，表甲那一页扔掉了十行；② **上下边距两片各补一整份**，不是「上归头片、下归尾片」；
③ **`w:trHeight` 每一片各要一份**，不是「整行算完把富余留给尾片」；
④ 猜对的那条是**头片照样认 `w:vAlign`**（原来一律按 top 摆，是错的）。
第三条顺手解释了另外两件本来要单独写规则的事：**一片都满足不了 `trHeight` 时整行挪走**
（表乙要 420pt、本页只剩 266pt，Word 挪了），以及**要的高度大过整页版心时续页顶上不重复表头**
（表乙的两片各占满一整页，两页顶上都没有表头；表丁的尾片只要 200pt，页顶就照常重复）。
第四问「接缝上画不画线」的答案是**画**，而且画的是**表级** `w:tblBorders` 的上下边 ——
不是这一行自己的（表己的第二行后面还跟着一行，它自己的下边框是 3pt 绿线，
接缝上画出来的却是 0.5pt 的黑外框）。也就是说 Word 把每一页上的表格片段**当成一张自己封口的表**画。
接缝线由布局层写进切片的 `cell.borders`（`seamBorders()`），**渲染层因此一个字都不用改**，
原来那个 `SPLIT_ROW_SEAM_BORDER` 开关连同它的猜测一起删了。尾片顶上那条线像正常格线一样占高度。
没做：嵌套表格的行不再往下切、格内不管孤行寡行。

**页眉页脚**（`packages/layout/src/header-footer.ts`）三条几何规则**都已实测**
（样本 `spike-header-01/02`，跑 `pnpm --filter @uw/fidelity spike:header`，落在 `HEADER_RULES`）：
① 页眉框顶 = `w:header`（到**纸**顶）；② 页脚量的是框**底**（框底 = 纸高 − `w:footer`）——
与页眉量顶边**不对称**，按对称写会让页脚整体偏一个页脚高度；③ **页边距是最小值不是固定值**：
版心顶 = max(`w:top`, 页眉底)、版心底 = min(纸高 − `w:bottom`, 页脚顶)。
直接后果是 **`PageGeometry` 每页一份**（首页页眉与偶数页页眉长度可以不同），
`availHeight()` 因此必须看**这一页自己的**版心 —— 原来的写法在 `breakPage()` 之后读的是
节的纸面几何，页眉进来之前两者恰好相等，这个洞才一直没露出来。
页脚里的 `{ PAGE }` **一趟就是准的**：它不走正文那张「run id → 文字」表（同一个 run 每页显示的
不是同一串字，那张表装不下），而走 `LayoutDocumentOptions.headerFields` 里的「怎么算」，
在开页那一刻算。**选择**规则（奇偶看显示页码、`titlePg` 没定义 first 时首页为空、跨节沿用、
补出来的空页算不算本节首页）按规范实现但**没有样本**，四问写在 `header-footer.ts` 的文件头。
三份样本（含 `spike-header-03`：首页 / 奇偶各一份 + 页脚里真的 `{ PAGE }`）进了 CI
（`header-fixture.test.ts`，跨平台，12 页逐行全对）。

**行高走哪一套规则也标定完了**（样本 `spike-script-01`，跑 `pnpm --filter @uw/fidelity spike:script`，
落在 `line-height.ts` 的 `SCRIPT_RULES` 与 `@uw/fonts` 的 `composeLineBox()`）。这一层原来有
一条**已知没做对**的：东亚字体里的**纯 ASCII 行**。答案是**看实际画字的那款字体**、**逐段**判 ——
不是「这一行有没有东亚字符」（旧实现），不是 `w:eastAsia` 槽，也不是 `w:hint`。
一行只有「A2C6」、ascii 槽里是等线，Word 照样按东亚规则给行高（20.32pt / 15pt 字，
拉丁规则只有 15.63pt）：**差 30%，而且每一行都差**。判据落在 `TextMeasurer.eastAsianFont()`
（查这款字体有没有 U+4E00 的字形）；字体缺失时它答 `undefined`，退回按字符判那条旧路 ——
谎报成拉丁字体会让一份缺字体的中文文档每行都矮 30%，比退回旧路错得远。
同一份样本顺带钉死了**混排行的合成**（原来是判断不是实测）：几款字体**各自的行盒逐项取 max**
（上取最高、下取最深），不是「取各自行高的最大值」—— 等线画 ASCII、宋体画汉字的那一行
Word 给 50.28pt，**比两款字体各自的行高都大**（48.77 / 46.80），旧写法按定义说不出这个数。
每款是在**自己的**自然行高里居中的，在合成后的行高里居中会把基线往下拽 0.63pt。
这与基线穿刺那条「东亚行的行盒只由东亚字体决定、拉丁 run 完全不参与」**不矛盾，是把它讲对了**：
等线 72pt + Times 72pt 那一页 Times 没赢，不是因为它没参与，而是因为它作为拉丁字体
走拉丁规则、核心盒上沿只有 67.22pt，输给等线的 69.57pt。两种说法在那一页上同解，
只有「两款东亚字体上下互不相让」的行分得开。排组合（8 种）逐页比，唯一满分 11/11（第二名 8/11），
样本进了 CI（`script-fixture.test.ts`）。**造样本时撞见一格 Word 自己造不出来的**：
`Font.NameFarEast = "Times New Roman"` 报 0x800A16D4，Word 界面里那个下拉框只列中日韩字体 ——
但与 `spike-table-03` 的相邻边框冲突不同，那一格不是唯一能回答问题的局面，所以没为它改 XML。

**分页的三条规则已经标定完**（样本 `spike-page-01/02`，跑 `pnpm --filter @uw/fidelity spike:page`，
落在 `page.ts` 的 `PAGINATION_RULES`）：孤行寡行**保底 2 行**、段前间距落在页首**不算**、
keepNext 的接缝要留出下一块**「最少能放多少」**（不是它的第一行，也不是整块）。
标定方式与别的 spike 不同 —— 分页规则不是一个数而是三条互相纠缠的判断，所以是把
3 × 2 × 3 种组合排开逐页比对，实现的这一组是**唯一**的满分（50 页全对，无并列）。
两份样本连同 gongwen-01 一起进了 CI（`page-fixture.test.ts`，跨平台）。

顺带被逼出来一条**行盒**的新结论：**固定值行距（`w:lineRule="exact"`）下基线 = 行高 × 0.8，
与字体、字号都无关**（`spike-baseline-04`，6 个样本比例 0.8002–0.8009）。原来的
「核心盒在行高里居中」是拿单倍 / 倍数 / 网格三种行距标定的，固定值那一格是空的 ——
`spike-page-01` 用固定行距 20pt，整页文字低了 1.77pt 才露出来。实现在
`@uw/fonts` 的 `baselineOffsetExact`，`line-height.ts` 按 `lineRule` 选哪一套。

**图片**横跨四层，每一层只解决一件事，中间只传一个 id：
① `@uw/model` 的 `parse-drawing.ts` —— **尺寸取 `wp:extent`**（用户拖出来的显示尺寸，
不是图片的像素尺寸）；`a:blip` 要**深搜**（规范路径 / `mc:AlternateContent` 的 Choice /
形状的填充三种写法），而图表与 SmartArt 里根本没有 blip，「找不到就是画不出来」正好是
画占位框的判据；裁剪 `a:srcRect` **不改外框**（裁剩的那块被拉伸回去）。VML（`w:pict`）
只取「图片引用 + `style` 里的外框」，其中不带单位的数字按 **px** 算（CSS 的默认单位，
不是 pt）—— 公文的红头与印章大量走 VML 这条路。
② `@uw/model` 的 `images.ts` —— 字节按**引用**收（包里常留着没人引用的 `media/image3.png`，
按类型全捞会把几 MB 的废弃扫描件读进内存），摊平成 `LoadedDocument.images`，
key = **部件前缀 + 关系 id**；前缀省不得：页眉里的 `rId1` 与正文里的 `rId1` 是两张图。
外链图（`r:link`）只给 URL，**不发网络请求**。
③ `@uw/layout` —— 内嵌图占宽占高，行盒三条规则已实测（见下段）；**带 `wp:anchor` 的图在文字流里占 0 宽**，
一个字都不许被它挤走，位置等整页排完再按 `wp:anchor` 的参照物换算成纸坐标
（`page.ts` 的 `placeFloats`，产出 `PageLayout.floats`）。**环绕方式不参与这个判断** ——
它回答的是「文字怎么让开」，没做的正是「让开」那一半（方形 / 上下型的对象位置与大小都对，
只是文字不绕着它走）。原来按 `wrap="none"` 判断、其余**退化成内嵌**，那是错的：
内嵌意味着它撑高所在的行，真实语料里页脚一个 144pt 的 `topAndBottom` 文本框
就这么把页脚撑到 145.9pt、再顺着「页边距是最小值」把版心挤掉 66pt，19 页排成 28 页。
④ `@uw/render-dom` —— `<image>` + 裁剪的 `clipPath` + 旋转翻转的 `transform`，
`href` 由宿主的 `imageHref(id)` 回调给（`imageHrefResolver(doc.images)` 是 data URI 版），
**渲染层不认识 OPC 包**。三处容易画错：`preserveAspectRatio="none"`（外框可以与图片不同比例）、
裁剪要「放大后再切」、旋转 90° 要把长宽比缩回去（extent 是转完的外接矩形）。
EMF / WMF 与图表画**尺寸正确的**虚线占位框、`alt` 进 `<title>` —— 框对了，周围的文字就不会错位。
结构性的那些由一份合成 docx 端到端兜着（`render-dom/src/image-docx.test.ts`，
含正文与页眉 `rId1` 撞车的用例）。未做：方形 / 上下型环绕（真的绕排）、表格单元格里的浮动对象。

**图片的几何也标定完了**（样本 `spike-image-01/02/03`，跑 `pnpm --filter @uw/fidelity spike:image`，
79 张图最大偏差 0.340pt、128 行最大偏差 0.120pt）。为它给真值管线加了一路新数据：`truth.json` 的 `pages[].images[]`，
从 PDF **算子表**里连着 CTM 演出来的图片落点 —— 图片在 PDF 里没有自己的坐标，
位置与大小全在矩阵里，`getTextContent()` 那一路根本看不见它。五条结论：
① **坐在基线上的是「盒」不是图**：盒高 = 图高**四舍五入到 1.5pt**（`objectBoxHeight`），
图在盒里靠上放，于是图底最多浮在基线以上 0.75pt。这条是被 0.1pt 步长的微阶梯逼出来的 ——
粗阶梯（只取偶数 pt）里它伪装成「h ≡ 4 (mod 6) 的那几档凭空多抬半磅」的噪声；
② **文字自己的下伸留着**：行高 = 盒高 + 文字下伸（仿宋 12pt 是 3.52pt、22pt 是 6.41pt，
跟着字号走）。**原来的实现让图把整行吃掉，每有一张图就少 3.5pt 并一路累积**，
这是这次标定照出来的最大的一个错；③ **`w:position` 对图片照样起作用**，行盒跟着变；
④ 浮动图的八种参照框（`page.ts` 的 `FLOAT_ORIGIN_RULES`）：两条与照规范猜的不一样 ——
**纵向的 inside/outside 镜像的是上下页边距**（不是版心，差着一整个上边距）、
**`character` 参照的是锚点前一个字**的左边缘。三条合起来落到数据上只剩一个数：
`LineObject.raise` = 对象底边高于基线多少，渲染层拿它一个就够，不必知道量化这回事。
⑤ **含图的行照样吸行网格**，且**倍数行距不乘在图撑起来的那一截上**（`spike-image-03`，
开着公文那套「每页 22 行」= 31.8pt）：吸的是「盒高 + 文字下伸」，吸到网格行的整数倍，
富余仍旧上下均分 —— 28pt 的图（32.02pt）吸成两行 63.6pt，60pt 那一档（63.52pt，比两行只矮
0.12pt）**仍是两行**，边界在 ceil 上不是四舍五入。倍数那一半是这份样本顺带照出来的**实现错误**：
网格 31.8pt + 1.5 倍 + 40pt 的图，按「合成一个自然行高再乘」（旧实现）得 95.4pt，Word 给 63.6pt。
实测的算法是**两侧分算**（`line-height.ts` 的 `advance()`）：文字侧「吸附 → 乘倍数」，
对象侧「对象要的高 + 倍数按**自然**行高多留的那段空白 → 吸附」，取大者作行的推进量；
基线在**赢的那一侧的行盒**里居中，而对象侧的行盒**不含**那段空白 ——
关掉网格的两档里图底严丝合缝坐在基线上（实测 40.53 vs 盒高 40.5），多留的 7.8pt 整个落在基线以下。
三份样本进了 CI（`layout/src/image-fixture.test.ts`），那里**行比的是逐行增量而不是绝对 y** ——
Word 自己的行位置带着 ±0.12pt 抖动，44 行图叠起来累到 1.5pt。

一份 docx 的入口是 `loadDocument(pkg, sink)`（`packages/model/src/load.ts`），产出
`body`（直接格式，可编辑）、`resolved`（级联完的纯数据，给布局）、`cascade`（上下文，**不可**过 Worker 边界）、
`fonts`、`numbering`、`fields`（配对好的域，单独一份而不是挂在树上 —— 域跨段落，挂上去就得补反向指针）、
`headerFooters`（**按关系 id 索引**的页眉页脚内容 —— 同一份常被多节共用，挂在节上要么复制几份、
要么变成跨节点引用；每份内容的节点 id 各带一个 `rId7:` 前缀，不带就会与正文的 `r0` 撞车，
页脚的页码会画到正文里去）。
`resolveBody()` 那一步就是 Worker 边界 —— `StyleSheet` 带方法，级联必须在过界前做完；
它同时是**编号计数器**跑的地方（编号「第几」只有按文档顺序走一遍才知道），结果落在
`ResolvedParaProps.numbering.label`（编号文字 + 它自己的字符属性 + `w:suff`）。
部件一律**按关系类型找**（`RelType.*`），不按 `word/styles.xml` 这种路径惯例猜；
页眉页脚是唯一的例外 —— 它们按**引用**（`sectPr` 里的 `r:id`）找，因为包里常留着没人引用的
旧 `header3.xml`，按类型全捞会把废弃的那些也画上去。

**原来卡在 Windows 的两件事都做完了**（基线穿刺 + 度量包抽取），Windows 侧现在只剩标定样本。
度量包：`packages/fonts/packs/*.json`，A/B/C/D 四类 17 款，**已入库**，所以消费侧完全跨平台 ——
`loadBundledPacks()` 一行注册进 `FontRegistry`，Mac / CI 上拿到的度量与 Word 用的一样。
重抽（换了字体或 Word 版本时）：`pnpm --filter @uw/fonts run packs`，非 Windows 上以退出码 2 拒绝跑。

直接后果：`layout/src/fixture.test.ts` 从「只能测与度量无关的性质」升级成**逐行比真值（L2）**，
现状是真实公文 18 行**对上 16 行**（行数一致、首末行一致）。这个数一路是
8 →（空格分桶 + 悬挂只吐空半边）11 →（临时挤压的三条规则 + 标点旁不加自动间距）16 涨上来的，
每一步都写在下面。剩下 2 行卡在一个至今解释不了的反例上。

顺着 L2 的差做完了**标点挤压**的标定（样本 `spike-punct-01`，跑 `pnpm --filter @uw/fidelity spike:punct`）：
**孤立的标点一点都不压，只有「标点紧跟标点」才压，固定 0.5 em** —— 这条推翻了从正文反推出来的猜测。
常态挤压落在 `items.ts` 的 `applyPunctPairs`（`buildItems` 阶段，给后一个标点一个**负的 `gapBefore`**，
与 Word 在 PDF 里的做法一致），塞不下时的临时挤压落在 `linebreak.ts` 的 `compress()`（挤**整行**的标点、
只挤到刚好够）。断行时三条补救的顺序是**先试着把标点整个塞进版心 → 悬挂 → 为多留一个字而挤压**，
与开发计划原先写的「压缩优先」相反（判据见下面两段）。

**悬挂标点**也已经按真值做完：吐出版心的**只是空的那半边**，墨留在版心内
（实测左边缘在版心线内 7.96pt、右边缘出界 8.05pt）。于是「能不能挂」看的是**半宽**塞不塞得下
（塞不下先挤压再挂），行宽把这半个字算进去（`break-class.ts` 的 `HANG_INSIDE_RATIO`）——
第 4 行那个悬挂的「，」右边缘与真值差 0.049pt。同一批真值还钉死了「算不算标点紧跟标点」
要看**接缝上有没有空白**：`「，`（开口紧跟收口）两边都是墨，**不压**；`】…` 压掉「】」的右半边，
**要压**（`punctPairCompressible`）。

**空格随邻居分桶**（同一天的第二条）：ASCII 空格按区段该进 ascii 桶，真值却说
「只要任一侧的邻居是东亚字，Word 就用东亚字体量它」（仿宋 0.5 em vs Times 0.25 em，
三号字差 4pt）。判断在 fonts（`neutralTakesEastAsia`），应用在 layout
（`items.ts` 的 `applySpaceFont`）—— 邻居**跨 run**，`splitFontRuns` 那一层看不见。

**宽度那一维也标定完了**（样本 `spike-width-01`，跑 `pnpm --filter @uw/fidelity spike:width`，
落在 `items.ts` 的 `WIDTH_RULES`）。144 种组合逐行跑，实现的这一组唯一满分 31/31（第二名 30/31）。
读数不靠反推：真值的 `TruthItem.font` **直接说出** Word 用哪款字体画了这个字 ——
字体一换，PDF 里就换一次 `Tf`、起一个新片段，所以「片段字体序列」就是分桶结果本身。
四条结论，两条是**真错**：
① **歧义字符**（EastAsianWidth = A：`§ ° ± × ÷ ·`…）**跟着 `w:hint` 走，与两侧邻居无关** ——
同样一段 `B§B°B±B`，hint=eastAsia 时 `§` 是宋体的 1 em、hint=default 时是 Times 的 0.5 em；
② 空格随邻居这条**与 `w:hint` 无关**（原来附带的「要 hint=eastAsia」是猜的），
且中性字符里**只有空格**这么走 —— `/` `-` 两种 hint 下都待在 ascii 桶里；
③ **中西文自动间距是 1/4 em 不是 1/8** —— 那个 1/8 是开发计划抄来的说法、从来没有过真值
（gongwen-01 的中西文之间本来就打了空格，量不到这个数），每一个中西文边界都少 4.5pt / 36pt 字；
④ 那 1/4 em 按**接缝前面那个字符**的字号算，不是东亚那一侧的（`中`36 + `E`12 + `中`36
实测 9.03 / 2.99pt，把字号反过来就换成 3.00 / 9.05pt）。
自动间距的例外表也补全了：**靠 hint 才进东亚桶的歧义字符旁边也不加**（与全角标点同理），
所以判据是**码点**（`isEastAsianCodePoint`）而不是分桶结果。
顺带照出一条**代码里的**错：`isLatinLetter` 把 Latin-1 的 `×`(U+00D7) 与 `÷`(U+00F7)
当成了字母（它们夹在字母区间中间），`中×中` 每个接缝平白多出 1/4 em。
两处方法论上的坑：**空格是唯一不能按字体名读的字符**（Word 画它时不换 `Tf`，
字体名跟着前一个字走、宽度才是另一款字体的，按字体名读会得出相反结论），
以及歧义字符**只能挑 Latin-1 段的** —— `※ ℃ ① Ⅰ` 这些 Times New Roman 有的没有，
字体回退会盖住分桶；度量包也只采样到 0xFF，之外的码点我们自己算不准。
`w:hint` 与 spike-table-03 的冲突边框一样**只能改 XML 写进去**，但拦路的东西不同：
那边是 Word **表达不了**那个局面，这边是 Word **不让你选** —— 对象模型里根本没有这个属性
（`patch-docx.ts` 的 `patchRunHints`）。

**临时挤压**（塞不下时额外再挤一点把字留住）也标定完了，靠两份专门的样本
（`spike-compress-01/02`，跑 `pnpm --filter @uw/fidelity spike:compress`）。做法是让每段
「恰好一行放不下最后一个字」，用**点为单位的右缩进**把可用宽一格格调窄，于是「要挤多少」连续可控。
三条结论，都在 `break-class.ts` / `linebreak.ts`：

1. **左对齐一格都不挤** —— 15 段全部换行。挤压是**两端对齐**才有的行为（`ctx.justified`），
   左对齐的右边本来就是毛边，挤出来的地方没有用处
2. 一个标点最多让出 **0.48 em**（`PUNCT_COMPRESS_MAX_EM`），不是 0.5 —— 与常态的相邻标点
   挤压（0.5 em）差的那 0.02 em 是实测差异，两者动的不是同一段空白
3. 挤到多少就宁可换行，是拿它跟**换行后要拉开的量**比出来的：
   `挤压量 × 字距数 ≤ 30.6 × 标点数 × 拉伸量`（`PUNCT_COMPRESS_STRETCH_K`）。
   七组阶梯（标点 1/2/3/4/6 × 行长 14/20/27 字）的翻转点全部落在预测的 ±0.1pt 内

顺带钉死的两条：**全角标点旁边不加中西文自动间距**（`（ascii`、`cs）` 实测间隙 0.05pt 以内 ——
标点自己带着空半边，再加就成了双份），以及**后置标点先试着整个塞进版心，塞不下才谈悬挂**
（真值第 13 行的「）」靠挤掉行内的「（」正好收进版心，然后换「，」挂出去）。

L2 剩下那 2 行（真值第 10 / 11 行）是**唯一一个解释不了的反例**：第 10 行 Word 只差 4.6pt
就能留住「出」、行内还有四个孤立标点给得起，却换了行。那一行的特别之处是有一串 16 个连着的标点、
12 个接缝已经在常态挤压里各压掉了半个字 —— 「已经压过的行还肯不肯再压」多半另有规则。
写在 `PUNCT_COMPRESS_STRETCH_K` 的注释里，**没有为它硬凑常数**。

`@uw/layout` 的用法：整份文档走 `layoutDocumentWithFields(resolved, fields, { measurer, settings, headerFooters })`
（不带域与页眉页脚时等价于 `layoutDocument(resolved, { measurer, settings })`）
→ `DocumentLayout`（`pages → blocks → lines/rows → fragments`，**有 y**）。
单块的两个入口仍然**不带 y**，它们是分页的输入、也是缓存的单位：
`layoutParagraph(resolvedParagraph, { measurer, contentWidth, settings, docGrid })`
→ `ParagraphLayout`（每行的 x / 逐字 x / 行高 / **行顶到基线** / 渲染片段）；
`layoutTable(resolvedTable, { …, availWidth })` → `TableLayout`（列宽 / 每格的 x 与可用宽 /
格内段落 / 每格四条边解析完的边框）。表格的列宽直接取 `w:tblGrid` ——
**Word 存盘时已经把 autofit 算完的结果写在那儿了**，照着用就与 Word 一致，
这也是「完整 autofit 算法」能列为非目标的原因。
未标定的常数一律集中在 `packages/layout/src/uncalibrated.ts`，每条都写了「拿什么样本能钉死」——
布局里出现别处的魔法数字视为 bug，散落的数字会被后人当成实测结论。

`@uw/fonts` 的用法：`FontRegistry` 收字体（`fontkitSource` 一级 / `metricsPackSource` 二级，
随库那 17 款走 `@uw/fonts/node` 的 `loadBundledPacks()`；
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

域那一层（`fields.ts`）有三处容易搞反：
① 域**跨段落**（TOC 能跨几十段），所以 `scanFields()` 先把整份 body 拉平成一条 run 流再配对 ——
按段落扫永远配不上，这也是它必须在 `resolveBody()` **之前**跑的原因；
② **没有 separate 的域什么都不显示**，不是「显示指令」；结果区里存的是 Word 上次算出来的文字，
直接显示就是「打开即所见」（求值见下，没有 separate 的域**也不求值**）；
③ 嵌套域的 `instrText` 归内层（靠栈分家），外层的指令文字因此**缺一块**（`IF { PAGE } = 1` 只看得到 `IF  = 1`），
回填是求值期的事。另外 `w:fldSimple`（压缩写法）走的是「压平成 run 上的标记」那条路，与 `w:hyperlink` 同理，
标记**带 id** —— 挨着的两个 `w:instr="PAGE"` 是两个域。`w:instrText` 与 `w:t` 相反，**不去首尾空白**：
去掉再拼会把 ` IF ` + ` = 1 ` 接成 `IF= 1`。

**域求值**在 layout 那一侧（`packages/layout/src/fields.ts`），不在 model —— 求值要页码，
页码是分页的产物。入口 `layoutDocumentWithFields(resolved, fields, opts)`，认
**PAGE / NUMPAGES / SECTIONPAGES**（TOC / SEQ 没做，照旧显示文件里存的旧结果）。四处要点：
① 结果**不写回模型**，而是外挂一张「run id → 显示的文字」的表当排版入参
（`LayoutDocumentOptions.fieldValues`）—— 写回去要每趟迭代克隆一棵树，模型里还会有两份真相；
② **收敛判据是那张表不再变**，不是「页数没变」：页数一样、某个 PAGE 从 3 变 4（内容在页之间挪位）
完全可能，架构 §6 原来写的判据是错的；
③ **没写 A→B→A 振荡检测**，开发计划 §2.4 说它「是必需品」，就这三个域而言说反了 ——
域文字只会变宽、分页规则只会把内容往后推，页数**单调不减**，回不了头。防线是
`MAX_FIELD_PASSES = 5`，撞上去按「取页数较大者冻结」退出 + 诊断。TOC（目录能变短）进来再补；
④ 没写 `\*` 的 PAGE 跟着**本节**的 `w:pgNumType w:fmt`（`SectionProps.pageNumFormat`，
顺手补进 model）—— 「前言罗马数字、正文阿拉伯数字」就是靠它。`\* CHINESENUM1|2|3`
的映射没有 Word 样本，关在 `uncalibrated.ts` 的 `FIELD_CHINESE_NUM_FORMATS`。
渲染出来的域结果带 `data-field="1"`：与编号相反，它**要**能被复制与 Ctrl+F 搜到，
但它不在 document.xml 里，反查不到 `DocPosition`。

表格那一层（`table-props.ts` / `parse-table-props.ts` / `cascade-table.ts`）也有四处容易搞反：
① 级联层序是「样式链自身属性 → 命中的条件格式（按 `CONDITIONAL_ORDER`）→ 直接格式」，
其中**行带排在列带之后、首末行排在首末列之后**（所以表头行会盖住首列的格式）；
② `w:tblLook` 是**开关**不是格式 —— 样式里定义了 `firstRow` 但 look 说不要，那份格式就不应用；
③ 表格样式的 `pPr` / `rPr` 铺给格内段落时排在**段落样式链之前**（走 `CascadeContext.tableStyleLayers`），
段落自己的样式要能盖掉表头行的加粗；
④ 单元格左右各 108 twips 的默认边距来自**默认表格样式**（`Normal Table`）而不是什么规范常数 ——
`w:tcMar` 缺席退到表级 `w:tblCellMar`，不是退到 0。
`w:tblPrEx`（**行级表格属性例外**，Word 从别处粘一行进来时写的，真实公文里很常见）盖在
**已级联完**的表级结果上（`applyRowExceptions`）—— 它是直接格式，插进样式链里反而要为它排新层序；
被它改过的表级边框经 `ResolvedRowProps.tableBorders` 带到布局层，
因为边框冲突解析的第一级「单元格没写就退到表级」对这一行说的是它。
隔行带的**序号算法**与 `CONDITIONAL_ORDER` 的**层序**已用 `spike-table-02` 标定完（见上），
其中层序原来是错的（列带盖行带，不是反过来）。

边框冲突解析（`layout/src/table-borders.ts`）是**两级模型，顺序不能反**：
① 层级覆盖 —— 单元格写了这条边就用它（**含 `w:val="nil"`**），没写才退到表级的
`top`/`insideH` 那一套；② 相邻竞争 —— 共享这条线的两个格子各出一个候选比大小。
`nil` 在 ① 里是强的（Word 里「擦掉某格的格线」就靠它），在 ② 里是弱的（一格 nil、
邻格 single 就画 single）；**合成一步会让整表的内部格线被一格的 nil 抹掉**。
另外水平边要**按列分段**（表头一格跨 3 列、下面 3 格，那条线分 3 段各比各的），
`vMerge=continue` 与上格之间不画线。**竞争规则已经实测**（`spike-table-03`，21 组 × 横竖两遍
42 条边全对，证据表在 `table-borders.ts` 的 `BORDER_CONFLICT_RULES`）：顺序是可见性 →
**线型分类**（点线 < 虚线 < 实线类，跨类时线宽一点都不算数：3pt 的点线输给 0.75pt 的单线）→
同为破折类就直接看位置（**同一种破折线之间连宽度都不比**）→ 实线类内部比**画出来的厚度**
（双线 = 3 × `w:sz`）→ 厚度打平比样式权重 → 仍平局取**左上**。原来照 CSS collapsing borders
类比写的「先比线宽、再比样式权重」前两条都是错的，后两条对。实测没覆盖的线型归哪一类、
算多厚仍在 `uncalibrated.ts`。

字体名有个坑：中文版 Word 写的是「黑体」「等线」这种本地化名，磁盘上的字体叫 `SimHei` / `DengXian`。
桥在 `fontTable.xml` 的 `w:altName`，查找顺序用 `fontNameCandidates()`，别只按一个名字查。

三份必读文档，动手前按需查：

| 文档 | 回答 |
|---|---|
| [docs/architecture.md](docs/architecture.md) | 代码怎么切、数据怎么流、坐标怎么管、五条设计原则 |
| [docs/api.md](docs/api.md) | 对外 API 长什么样 |
| [docs/DEVELOPMENT-PLAN.md](docs/DEVELOPMENT-PLAN.md) | 阶段顺序、每阶段 DoD、非目标清单、**每做完一件事的收尾动作（§8）** |
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
pnpm --filter @uw/playground dev     # 调试台，:5273（拖一份 docx 进去就画）

# 单个包 / 单个测试文件 / 单个用例
pnpm --filter @uw/fonts run test
pnpm --filter @uw/fonts run test src/metrics.test.ts
pnpm --filter @uw/fonts run test src/metrics.test.ts -t "东亚"

# 真值流水线（仅 Windows + 已装 Word）
pnpm truth                           # 只重算过期 fixture
pnpm truth gongwen-01                # 指定 fixture
pnpm truth --force
pnpm --filter @uw/fidelity spike           # Phase 0 行高穿刺
pnpm --filter @uw/fidelity spike:baseline  # 基线穿刺（基线在行高里的位置）
pnpm --filter @uw/fidelity spike:punct     # 标点挤压穿刺（什么时候压、压多少）
pnpm --filter @uw/fidelity spike:compress  # 临时挤压穿刺（塞不下时肯挤多少才换行）
pnpm --filter @uw/fidelity spike:page      # 分页穿刺（孤行寡行 / keepNext / 页首段前间距，**不需要 Word**）
pnpm --filter @uw/fidelity spike:header    # 页眉页脚穿刺（框摆在哪 / 怎么挤版心，**不需要 Word**）
pnpm --filter @uw/fidelity spike:image     # 图片穿刺（内嵌图的行盒 / 浮动图的参照框 / 行网格与倍数行距，**不需要 Word**）
pnpm --filter @uw/fidelity spike:table     # 表格穿刺（格线占不占高 / 吃不吃宽 + 条件格式的命中与层序，**不需要 Word**）
pnpm --filter @uw/fidelity spike:table-border  # 格线冲突穿刺（相邻两格各写一条边，Word 画哪一条，**不需要 Word**）
pnpm --filter @uw/fidelity spike:table-split   # 拆行穿刺（切在哪一页 / 边距怎么分 / trHeight 归谁 / 头片 vAlign，**不需要 Word**）
pnpm --filter @uw/fidelity spike:script    # 脚本穿刺（纯 ASCII 的行走东亚还是拉丁规则 / 混排行怎么合成，**不需要 Word**）
pnpm --filter @uw/fidelity spike:width     # 宽度穿刺（歧义字符与空格进哪个桶 / 中西文自动间距多宽，**不需要 Word**）
pnpm --filter @uw/fidelity preview -- --truth  # fixture 画成 out/*.html 并叠真值基线（不需要 Word）
pnpm --filter @uw/fidelity corpus          # 语料体检：整份文档与 truth.json 比页数 / 每页行数 / 每行文字 + 汇总诊断
pnpm --filter @uw/fidelity corpus 名字 -- --diff   # 逐行看差异
pnpm --filter @uw/fonts run packs          # 重抽度量包（仅 Windows，产物入库）
pnpm --filter @uw/fonts run packs:check    # 只校验入库的包与本机字体是否一致
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

`apps/fidelity` 用 Word COM 导出 PDF、pdf.js 抽每个文本片段的 transform（外加算子表里的
**图片落点** `pages[].images[]` 与**画出来的线** `pages[].rules[]`），产出**坐标级**真值
`fixtures/*.truth.json`（单位 pt，原点页面左上角、y 向下、`y` 是基线）。这不是测试工具，是架构的一部分：
它直接约束 `LayoutResult` 的数据形状 —— 必须能逐行、逐片段与真值 diff。

断言分级随阶段收紧：L0 页数 → L1 每页首末行文本 → L2 每行断行点 → L3 基线 y 误差 < 0.5pt → L4 片段起始 x 误差 < 0.5pt。

写布局代码时的判据是「与真值差多少 pt」，不是「看起来像不像」。改 `metrics.ts` 之类的标定结果前先看
`metrics.test.ts` —— 里面的期望值是从 Word 实测反推的，真值站在测试那一边。

## 已标定与未决

**已定死（Phase 0，13 个样本最大误差 0.132pt）** —— 实现在 [`packages/fonts/src/metrics.ts`](packages/fonts/src/metrics.ts)：

| 这段文字用的字体 | 单倍行距行高 |
|---|---|
| **东亚字体** | `(usWinAscent + usWinDescent) × 1.3 × 字号 / unitsPerEm`，**不加**外部行距 |
| **拉丁字体** | `(usWinAscent + usWinDescent + GDI 外部行距) × 字号 / unitsPerEm` |

那个 1.3 是乘在**字体度量**上，不是「1.3 × 字号」；只测宋体家族（unitsPerEm=256、win 跨度恰好 1.0em）两种假设分不开。

⚠️ 表头这一列原先写的是「**含东亚文字**的行 / **纯拉丁文字**的行」，
**被 `spike-script-01` 推翻了**：判据是**字体**不是字符，而且**逐段**判（见下面的脚本穿刺）。
Phase 0 这 13 个样本分不开两种说法 —— 纯拉丁的行用的是 Times New Roman，
「字符是拉丁的」与「字体是拉丁字体」在它们身上完全重合。

**已定死（基线穿刺，30 个样本最大误差 0.140pt）** —— 实现在同一个文件的 `baselineOffset()`：
**「核心盒」在最终行高里居中**，核心盒 = win 跨度（拉丁的话再加上 GDI 外部行距）。
推论是行距倍数、网格吸附拉出来或压掉的空间**一律上下均分**，不用为每种来源各写一条规则。

⚠️ 这条推论原先还写着「固定值行距」，**被 `spike-baseline-04` 推翻了**：
固定值行距（`w:lineRule="exact"`）下基线 = **行高 × 0.8**，与字体、字号都无关
（`baselineOffsetExact`）。原来会写错，是因为前三份 fixture 的 `lineSpacingPt` 全是 0 ——
「固定值」那一格从来没测过，是顺着「多出来的空间一律均分」推出去的。
露馅的是分页样本 `spike-page-01`（固定行距 20pt）：整页文字比预测低 1.77pt，
正好是仿宋 12pt 那 30% 额外行距的一半。

同一批 fixture（`spike-baseline-01|02|03`，跑 `pnpm --filter @uw/fidelity spike:baseline`）
还顺手钉死了三条容易搞反的：
① **网格吸附在行距倍数之前** —— 网格 31.8pt 开 1.5 倍行距，仿宋 16pt 与宋体 12pt 的行高都是 47.7pt，与字号无关；
② **东亚行的行盒只由东亚字体决定**，拉丁 run 完全不参与（判据是等线 72pt + Times 72pt 那一页：
Times 的 winAscent 更大，若参与就该赢，实测却仍是等线单独的值）；
③ **空段落走段落标记的 ascii 桶 + 拉丁规则** —— 12pt 空段落的行高是 13.78pt（Times 的 1.1499 em）
而不是宋体的 15.6pt，因为段落标记本身不是东亚字符。

**已定死（分页穿刺，两份样本 50 页全对）** —— 实现在
[`packages/layout/src/page.ts`](packages/layout/src/page.ts) 的 `PAGINATION_RULES`：
孤行寡行保底 **2 行**、段前间距落在页首**不算**、keepNext 的接缝留出下一块**「最少能放多少」**。
判据不是残差而是「哪一组能逐页复现 Word」：3 × 2 × 3 种组合排开跑，实现的这组唯一满分。

**已定死（图片穿刺，三份样本 79 张图、128 行，最大偏差 0.340 / 0.120pt）** —— 实现在
[`packages/layout/src/line-height.ts`](packages/layout/src/line-height.ts) 的 `OBJECT_RULES`
与 [`packages/layout/src/page.ts`](packages/layout/src/page.ts) 的 `FLOAT_ORIGIN_RULES`：
内嵌图占的高度 = 图高**四舍五入到 1.5pt**（坐在基线上的是这个盒，图在盒里靠上放）、
文字的下伸留着（行高 = 盒高 + 文字下伸）、`w:position` 对图片起作用、
**含图的行照样吸行网格**但**倍数行距不乘在图撑起来的那一截上**（两侧分算再取大）；
浮动图纵向的 inside/outside 镜像**上下页边距**、`character` 参照锚点**前一个**字。
判据同样是排组合（24 种）逐行逐图比，唯一满分。**「图的底边坐在基线上」这句原话是错的** ——
它在偶数 pt 的粗阶梯上看着成立（偏差被伪装成噪声），0.1pt 步长的微阶梯才照出那是个台阶。

**已定死（页眉页脚穿刺，三份样本 12 页全对）** —— 实现在
[`packages/layout/src/header-footer.ts`](packages/layout/src/header-footer.ts) 的 `HEADER_RULES`：
页眉框顶 = `w:header`、页脚量的是框**底**、**页边距是最小值**（页眉页脚长过它就把版心顶开）。
同样是排组合（8 种）逐页比，唯一满分。**页脚量底边这一条最容易搞反** ——
量顶边的那四组在三份样本上全军覆没（每份都差一个页脚高度）。

**已定死（表格穿刺，两份样本 123 段里 122 段对上）** —— 实现在
[`packages/layout/src/table.ts`](packages/layout/src/table.ts) 的 `TABLE_RULES` 与
[`packages/model/src/table-props.ts`](packages/model/src/table-props.ts) 的 `CONDITIONAL_ORDER`：
**水平格线占整条线的宽（纵向）、竖格线不吃可用宽**、条件格式里**列带盖行带**而**首末行盖首末列**。
排组合（3 × 2）逐段比，唯一满分；差的那一段是 `w:tcMar left="0"` 那一格的 0.59pt，没有模型（见上）。

**已定死（格线冲突穿刺，一份样本 21 组 × 横竖两遍 42 条边全对）** —— 实现在
[`packages/layout/src/table-borders.ts`](packages/layout/src/table-borders.ts) 的
`BORDER_CONFLICT_RULES`：**先按线型分类**（点线 < 虚线 < 实线类，跨类时线宽完全不算数）、
**同一种破折线之间不比宽度**（直接看位置）、实线类内部比**画出来的厚度**（双线 = 3 × `w:sz`）、
厚度打平比样式权重、仍平局取**左上**。判据同样是排组合（32 种）逐边比，唯一满分。
读数靠**颜色**：竞争两侧各给一个独一无二的颜色，画出来的线是什么颜色就是谁赢 ——
与条件格式那份拿字号认层序同一招。**Word 自己造不出这个局面**（对象模型里一条共享边
只有一个 Border 对象），所以冲突是改 XML 写进去的，见 `apps/fidelity/src/patch-docx.ts`。

**已定死（拆行穿刺，一份样本 14 页逐行全对）** —— 实现在
[`packages/layout/src/table-split.ts`](packages/layout/src/table-split.ts) 的
`TABLE_SPLIT_RULES` 与 [`packages/layout/src/page.ts`](packages/layout/src/page.ts) 的 `placeTable()`：
**就地切**、**上下边距两片各补一整份**、**`w:trHeight` 每一片各要一份**、
**头片照样认 `w:vAlign`**；接缝上那两条线**画**，取的是表级 `w:tblBorders` 的上下边。
排组合（16 种）逐页比，唯一满分。四条里**三条原来是反的**，而且每一条都改分页 ——
「挪到下一页再切」那一版整份样本 14 页里只对上 3 页。
判页的方式与别处不同：Word 自己的行距带着 ±0.12pt 抖动、十八行能累到 0.58pt，
所以比的是**首行绝对 y + 逐行增量**（与 `spike-image` 同一个理由）。

**已定死（脚本穿刺，一份样本 11 页逐行全对）** —— 实现在
[`packages/layout/src/line-height.ts`](packages/layout/src/line-height.ts) 的 `SCRIPT_RULES` 与
[`packages/fonts/src/metrics.ts`](packages/fonts/src/metrics.ts) 的 `composeLineBox()`：
**行高走东亚规则还是拉丁规则，看的是「实际画字的那款字体是不是东亚字体」，逐段判** ——
不是行里有没有东亚**字符**（旧实现）、不是 `w:eastAsia` 槽、也不是 `w:hint`；
**同一行里几款字体，各自的行盒逐项取 max**（上取最高、下取最深），不是「取各自行高的最大值」。
排组合（8 种）逐页比，唯一满分 11/11（第二名 8/11）。判据放在 `TextMeasurer.eastAsianFont()`
（查这款字体有没有 U+4E00 的字形），字体缺失时答 undefined、退回按字符判那条旧路。

**已定死（宽度穿刺，一份样本 31 行逐行全对，行末 x 最大偏差 0.089pt）** —— 实现在
[`packages/layout/src/items.ts`](packages/layout/src/items.ts) 的 `WIDTH_RULES` 与
[`packages/fonts/src/script.ts`](packages/fonts/src/script.ts) 的 `bucketOf` /
`neutralTakesEastAsia`：**歧义字符跟着 `w:hint` 走、与邻居无关**；中性字符里
**只有空格**随东亚邻居（任一侧就算）且**与 hint 无关**；**中西文自动间距 = 1/4 em**
（不是一直写着的 1/8），按**接缝前面**那个字符的字号算；靠 hint 才进东亚桶的歧义字符
旁边**不加**间距。144 种组合逐行比，唯一满分。读数是**字体名**不是宽度
（`TruthItem.font` 直接说出 Word 用哪款字体画了这个字），
唯一的例外是空格 —— Word 画它时不换 `Tf`。

**未决**：同一行里两款东亚字体且**字号也不同**时怎么合成 —— `spike-script-01`
只覆盖了同字号那一格。样本这一侧的限制 2026-09-05 解开了：`make-fixture.ps1`
认 `runs`（盖在 `text` 上的一层，段内逐段改字体 / 字号），造这份样本不再需要改工具。

行高与基线之外还有七处未标定（临时挤压那条原来排第一，已经用 `spike-compress-01/02` 做完，见上）。
一原先是 `splitFontRuns()` 的歧义字符集与空格分桶的两处边界，2026-09-05 用 `spike-width-01`
做完了（**歧义字符跟 `w:hint`、空格与 hint 无关、`/` `-` 不随**，另外照出中西文间距
是 1/4 em 不是 1/8，见上）。当时估的「拿 `TruthItem.font` 直接读」是对的，
但那一句里列的候选字 `α β Б “ ”` **用不得**：度量包只采样 0x20–0x7E 与 0xA0–0xFF，
之外的码点我们自己一律退到 `defaultAdvance`，Word 答得出、我们算不准 ——
真正能用的只有 Latin-1 那一段。这一格现在空着。
二是编号的三个样本：`w:lvlJc="right"` 的编号以哪条线对齐、编号宽过悬挂缩进时正文落在哪、
`chineseCounting` 与 `chineseCountingThousand` 在 105 / 1005 上各显示什么。
三原先是**表格隔行带的序号**，2026-08-26 用 `spike-table-02` 做完了（**归属与实现一致，
层序错了一条**，见上）。它留下的「东亚字体里的纯 ASCII 行走哪套行高规则」2026-09-05
也用 `spike-script-01` 做完了（**看字体**，见上），那份样本还顺带把「混排行怎么合成」
一起钉死了 —— 它本来是第一条那份样本的一半，另一半（**宽度**：歧义字符与中性字符进哪个桶）
2026-09-05 也用 `spike-width-01` 做完了（见上）。
四原先是**表格边框冲突**，2026-08-27 用 `spike-table-03` 做完了（见上）。当时估的
「一张 2×2 的表就够」偏小 —— 光是「宽度与样式谁先比」就要五组互为镜像的配对才排除得掉
别的解释，最后用了 21 组；更没料到的是**样本造不出来**：Word 的对象模型里一条共享边
只有一个 Border 对象，设一侧等于两侧都设，冲突只能改 XML 写进去
（`apps/fidelity/src/patch-docx.ts`）。它留下的新问题是**实测没覆盖的那些线型**：
`triple` 算几倍厚、`dashSmallGap` / `dotDash` / `wave` 各归哪一类（现在照样子归的，
在 `uncalibrated.ts`）—— 照着已有的组再加几对配对就能钉死。
五是域的 `\* CHINESENUM1|2|3` 各对应哪套中文数字（`FIELD_CHINESE_NUM_FORMATS`）：
一份三节的 docx，页码跑到 11 与 105，看 Word 显示「十一 / 拾壹 / 一〇五」里的哪一种。
中文数字比阿拉伯数字宽，猜错了页码那一行的断行点也会跟着错。

六原先是**图片参不参与行网格吸附**，2026-08-26 用 `spike-image-03` 做完了（**参与**，见上）。
这一格前后作废过三条：`OBJECT_SITS_ON_BASELINE`（说法本身就是错的，坐在基线上的是盒不是图）、
`FLOAT_RELATIVE_FROM_CALIBRATED`、以及刚做完的网格这一条。它留下的新问题只有一个：
**倍数行距小于 1 时图会不会被压扁** —— `advance()` 把「倍数多留的空白」夹到了 0，
按「对象要的高是硬下限」处理，没有样本；一份「0.5 倍行距 + 一张比行高的图」就能钉死。

七是页眉页脚的**选择**规则四问（几何三条已实测，选择这四条只有规范做依据）：奇偶看的是显示页码
还是物理页序、`w:titlePg` 开着却没定义 first 时首页是不是空的、本节没写某一类时沿不沿用上一节、
`evenPage` 补出来的空页算不算「本节首页」。四问都不改坐标，只改「画的是哪一份」，
写在 `header-footer.ts` 的文件头，每问都带着能钉死它的样本。

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

## 做完一件事之后（开发计划 §8 的摘要）

1. `pnpm turbo run typecheck test` + `pnpm lint` 全绿；真值闸门（`MIN_L2_MATCH`）**只许往上调**
2. 新量到的数字连证据表一起写进注释；未标定的进 `uncalibrated.ts` 并写清「拿什么样本能钉死」
3. **回头对照 [docs/architecture.md](docs/architecture.md) 与 [docs/api.md](docs/api.md)，
   把被真值 / 代码推翻的表述改对**，并把「原来为什么是错的」一起留下 ——
   这两份写在实现之前，实现却是被真值推着走的，它们**必然**会过期
4. 更新进度三处：本文件的「当前进度」段、开发计划对应 Phase、架构 §3.2 现状表
5. 提交信息写清楚「为什么」

## 非目标（不要为它们预留扩展点）

紧密型/穿越型环绕、多栏排版、修订痕迹的**编辑**（只做显示）、OMML 数学公式排版、
VML / SmartArt / 图表（降级占位图）、`.doc` 二进制格式、完整 autofit 表格算法、RTL 与复杂文字。
需要时再改架构，比现在摆一堆空接口便宜。
