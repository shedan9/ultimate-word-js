/**
 * 未标定的常数 —— 全项目**唯一**允许出现「拍脑袋数字」的地方。
 *
 * 规矩（原则 1.5 的推论）：布局里任何一个不是从 Word 真值反推出来的系数，都必须放进这个文件，
 * 并写清楚「拿什么样本能把它钉死」。散落在各处的魔法数字会被后人当成实测结论，
 * 那比留一个洞危险得多 —— metrics.ts 里的 1.3 之所以可信，正因为它旁边贴着 13 个样本的误差表。
 *
 * 这里的值影响的是**宽度**，也就是断行位置。行盒与基线（`@uw/fonts` 的 `lineMetrics` /
 * `baselineOffset`）、标点挤压的三条规则（`break-class.ts` 的 `PUNCT_PAIR_COMPRESS_EM` /
 * `PUNCT_COMPRESS_MAX_EM` / `PUNCT_COMPRESS_STRETCH_K`）都已经实测标定，搬去了实现文件旁边 ——
 * **这个文件只留还没有真值的那些** —— 分页那三条（孤行寡行下限、页首段前间距、keepNext 的接缝）
 * 曾经在这里待过一天，`spike-page-01/02` 一跑就搬去了 `page.ts` 的 `PAGINATION_RULES`；
 * 图片的两条（内嵌图坐在哪、浮动图的六种参照物）也一样，`spike-image-01/02` 一跑就搬去了
 * `line-height.ts` 的 `OBJECT_RULES` 与 `page.ts` 的 `FLOAT_ORIGIN_RULES`；
 * 中西文自动间距（`AUTO_SPACE_EM`）与歧义字符 / 中性字符的分桶是 `spike-width-01`，
 * 搬去了 `items.ts` 的 `WIDTH_RULES` —— 那一条留在这里的两年里一直写着 1/8 em，
 * **实测是 1/4**，也就是每一个中西文边界都少算一半。`layout/src/fixture.test.ts` 现在 18 行对上 16 行，
 * 剩下 2 行是一个至今解释不了的反例，写在 `PUNCT_COMPRESS_STRETCH_K` 的注释里。
 * 末尾那三张**边框线型**的表是个例外：它们影响的是**画哪条线**，不改坐标。
 * 冲突规则本身已经用 `spike-table-03` 实测完了（见 `table-borders.ts`），
 * 留在这里的只剩「实测没覆盖到的那些线型归哪一类、算多厚」。
 */
import type { Twips } from '@uw/core';

/**
 * 小型大写字母（`w:smallCaps`）里，小写字母实际用多大的字号。
 *
 * 钉死办法：一段 `w:smallCaps` 的英文，量首字符到末字符的总宽，与按此系数的预测对比。
 * 公文语料里几乎不出现，优先级最低。
 */
export const SMALL_CAPS_SCALE = 0.8;

/**
 * 上下标（`w:vertAlign`）的字号系数。**只影响宽度**，基线升降量是另一回事
 * （那个连同行盒一起卡在基线穿刺）。
 *
 * 钉死办法：`x²` 这样的样本，量上标字符的推进宽度 ÷ 正文同字符的推进宽度。
 */
export const VERT_ALIGN_SCALE = 2 / 3;

/**
 * `w:*Chars`（1/100 字符）里「一个字符」的宽度 = 字号 × 这个系数。
 *
 * 取 1.0：中文字体的汉字推进宽度精确等于 1 em，公文的「首行缩进 2 字符」正是 2 个汉字宽。
 * 不去实际测一个汉字的宽度，是因为段落可能一个汉字都没有（纯英文段也能写 firstLineChars），
 * 那时按字号算才是 Word 的行为。
 */
export const CHAR_UNIT_EM = 1.0;

/**
 * 格内文字在单元格左边缘之外**额外**的一点右移，twips。
 *
 * `spike-table-01` 量出来的**残差**，不是推出来的规则：Word 把格内文字放在
 * 「格左边 + `w:tcMar`」再往右一点点，而那一点点跟边距对不上号 ——
 *
 * | 那一格的 `w:tcMar` 左边距 | Word 比「格边 + 边距」多出 |
 * |---|---|
 * | 5.4pt（默认的 108 twips） | 0.32pt |
 * | 20pt | 0.24pt |
 * | 0 | **0.59pt** |
 *
 * 三个数凑不出一条规则（不是常数、不是比例、也不是「至少多少」加常数），
 * 所以这里留 0，**不为它硬凑**。前两档都在 L4 判据（0.5pt）以内，只有零边距那一格
 * 越了线 —— 那也是 `spike:table` 唯一对不上的一项。
 *
 * 钉死办法：一张表，同一列上把 `w:tcMar` 的左边距排成 0 / 1 / 2 / 4 / 8 / 16 / 32pt
 * 七级阶梯，再复制一份把边框宽度从 0.5 换成 4pt。两份一比，就知道那一点点是跟着边距走、
 * 跟着边框走，还是一个下限。
 */
export const TABLE_CELL_TEXT_INSET: Twips = 0;

/** 把「em 的倍数」换成 twips。字号本身就是 twips，所以是一次乘法 */
export function em(fontSize: Twips, ratio: number): Twips {
  return fontSize * ratio;
}

/**
 * 边框冲突里每种线型算**哪一类**（大的赢）。
 *
 * 三档由 `spike-table-03` 实测（证据表在 `table-borders.ts` 的 `BORDER_CONFLICT_RULES`）：
 * 点线 0 < 虚线 1 < 实线类 2，而且**跨档时线宽一点都不管用** —— 3pt 的点线输给 0.75pt 的
 * 单实线、3pt 的虚线输给 0.5pt 的单实线。实测覆盖的只有 `single` / `double` / `dashed` /
 * `dotted` 四种，其余线型是**照样子归类的**，所以这张表关在这里。
 *
 * 认不出的 `w:val` 落到实线那一档：实文件里的生僻线型（`dashDotStroked`
 * `thickThinMediumGap` …）都是实线的花样，退成点线会让它凭空输掉。
 *
 * 钉死办法：`spike-table-03` 的 spec 里照着 dashed / dotted 那两组再加几组，
 * 把 `dashSmallGap` / `dotDash` / `wave` / `triple` 各与 single 配一次即可。
 */
const BORDER_STYLE_CLASS: Record<string, number> = {
  // 点线（实测）
  dotted: 0,
  // 虚线族（`dashed` 实测，其余照样子归类）
  dashed: 1,
  dashSmallGap: 1,
  dotDash: 1,
  dotDotDash: 1,
  // 实线类（`single` / `double` 实测，其余照样子归类）
  single: 2,
  thick: 2,
  wave: 2,
  double: 2,
  triple: 2,
  doubleWave: 2,
  dashDotStroked: 2,
  thickThinSmallGap: 2,
  thinThickSmallGap: 2,
  thickThinMediumGap: 2,
  thinThickMediumGap: 2,
  thickThinLargeGap: 2,
  thinThickLargeGap: 2,
  thinThickThinSmallGap: 2,
  thinThickThinMediumGap: 2,
  thinThickThinLargeGap: 2,
};

/** 0 点线 / 1 虚线 / 2 实线类。认不出的算实线类 */
export function borderStyleClass(style: string): number {
  return BORDER_STYLE_CLASS[style] ?? 2;
}

/**
 * 实线类内部比的是**画出来有多厚**，这张表是「`w:sz` 的几倍」。
 *
 * `double` 的 3 倍是实测的（`spike-table-03`：sz=12 的双线画出来 4.32pt = 3 × 1.44pt，
 * 于是它赢过 3pt 的单线、输给…… 见 `BORDER_CONFLICT_RULES` 的证据表）。
 * 多线型的其余几种按「几条线加几个缝」推：triple 5 倍，thick 与 single 同宽。
 * 破折类不走这条路（它们之间根本不比宽度），填 1 只是为了这张表是全的。
 *
 * 钉死办法：与上面那张表同一份样本 —— 让 triple 与 single 配一次，
 * 挑一对「按 5 倍算谁赢、按 3 倍算换一个人赢」的宽度就分得开。
 */
const BORDER_THICKNESS_FACTOR: Record<string, number> = {
  double: 3,
  doubleWave: 3,
  triple: 5,
  thickThinSmallGap: 3,
  thinThickSmallGap: 3,
  thickThinMediumGap: 3,
  thinThickMediumGap: 3,
  thickThinLargeGap: 3,
  thinThickLargeGap: 3,
  thinThickThinSmallGap: 5,
  thinThickThinMediumGap: 5,
  thinThickThinLargeGap: 5,
};

/** 一条边画出来的厚度 = `w:sz` × 这个系数。认不出的按 1 倍算 */
export function borderThicknessFactor(style: string): number {
  return BORDER_THICKNESS_FACTOR[style] ?? 1;
}

/**
 * 厚度打平之后的样式权重（大的赢）。实测只钉死了一处：**同样画出来 1.44pt 厚**时，
 * 0.5pt 的双线赢过 1.5pt 的单线（`spike-table-03` 的第「癸」组）。
 * 于是这里只分两档「多线型 > 单线型」，多线型之间谁赢没有样本。
 */
export function borderSolidRank(style: string): number {
  return borderThicknessFactor(style) > 1 ? 1 : 0;
}

/**
 * `\* CHINESENUMn` 三个开关各对应哪套中文数字。
 *
 * 前面几个数字格式开关（`\* Arabic` / `\* roman` / `\* alphabetic` / `\* Ordinal`）是
 * Word 域参考里写死的，算不上「拍脑袋」，所以留在 `fields.ts`；中文这三个**只有中文版
 * Word 帮助里的一句话**，没有样本，所以关在这里：CHINESENUM1 按「一、二、十一」，
 * CHINESENUM2 按法定大写「壹、贰」，CHINESENUM3 按「〇一二」逐位念。
 *
 * 猜错的代价不只是难看：中文数字比阿拉伯数字宽，页码宽度变了断行点就可能变。
 *
 * 钉死办法：一份三节的 docx，页眉里分别放 `{ PAGE \* CHINESENUM1|2|3 }`，
 * 页码跑到 11 与 105，看 Word 显示的是「十一 / 拾壹 / 一〇五」里的哪一种。
 * （其中 `ideographDigital` 自己的读法也还没标定，见 `@uw/model` 的 number-format.ts）
 */
export const FIELD_CHINESE_NUM_FORMATS: Readonly<Record<string, string>> = {
  chinesenum1: 'chineseCounting',
  chinesenum2: 'chineseLegalSimplified',
  chinesenum3: 'ideographDigital',
};

/**
 * 段落边框（`w:pBdr`）与底纹（`w:pPr/w:shd`）的几何 —— **整块没有真值**，本机没有 Word。
 * 按 Word 界面行为与规范写的几条，每条都改坐标（除了最后一条）：
 *
 * 1. **让出的高度 = 线的厚度 + `w:space`**，上边框让在段前间距与首行之间、下边框让在末行与段后间距之间
 *    （段前段后间距在框**外**）。厚度沿用表格格线那张 `BORDER_THICKNESS_FACTOR`（双线 3 倍）
 * 2. **成组**：相邻两段的可见边框、底纹、框的左右边（即左右缩进）全都一样才框在一起 ——
 *    就是下面这张键表。组内段与段之间的段后 + 段前间距在框**里**，底纹与竖线接过去
 * 3. 组内相邻两段之间画 `between`，它像上边框一样让在**下面那段**的段前间距之后；没有就什么都不让
 * 4. 框的左边取首行与其余行里更靠左的那个（悬挂缩进的编号在框里），竖线画在框外 `space` 处，
 *    **不挤文字区**
 * 5. 行网格：边框让出的高度不吸网格（公文开着网格时，带边框的段落之后每行都可能偏）
 *
 * 钉死办法：一份关掉网格的 docx，几段分别带「只有上 / 只有下 / 四周 / 双线 / 不同 space」的边框，
 * 再来三段一样的与一段缩进不同的（成组），用真值量每段首行基线与上一段末行基线的差 ——
 * 与 `spike-table-01` 量格线占高同一招；线本身的位置读 `truth.json` 的 `pages[].rules[]`。
 * 跨页时框封不封口是画法的事，见 render-dom 的 `uncalibrated.ts`。
 */
export const PARA_FRAME_GROUP_KEYS = ['borders', 'shading', 'x', 'width'] as const;

/**
 * `SEQ … \s N`（「遇到第 N 级标题就从头数」，中文报告的「图 2-1」就是 `STYLEREF 1` + `SEQ 图 \s 1`）
 * 认哪些标题 —— **只认第 N 级**，还是**第 N 级及更高**（大纲级别数值更小）的都算。
 *
 * 取「及更高」：`\s 2` 的题注遇到新的一章（标题 1）却接着上一章的号往下数，
 * 章节号已经换了、序号却没归零，与「按节编号」的本意相悖。Word 域参考只说
 * 「resets the sequence number to the heading level following the s」，没说更高级的标题算不算。
 * 标题认的是**级联后的大纲级别**（`outlineLevel` 0–8，样式里的 `w:outlineLvl` 也算），
 * 不是样式名 —— 中文版 Word 的「标题 1」样式 id 是 `1`，按名字认会漏。
 *
 * 钉死办法：一份 docx，标题 1 → 标题 2 → 两个 `{ SEQ 图 \s 2 }` → 标题 1（中间**没有**标题 2）
 * → 一个 `{ SEQ 图 \s 2 }`，全选 F9 更新后看最后那个是 1 还是 3。不改坐标，只改显示的数字
 * （数字位数变了才会连带改断行）。
 */
export const SEQ_RESET_ON_HIGHER_HEADINGS = true;

/**
 * `STYLEREF … \s` / `\t` 从编号里留下哪些字符。Word 域参考对 `\t` 的原话是「去掉非分隔符、
 * 非数字的文字」，`\s` 没写进参考（Word 插入带章节号的题注时写的就是 `STYLEREF 1 \s`），
 * 按中文用户最常遇到的现象当成同一件事：章节编号「第1章」的题注显示「图 1-1」、
 * 「第一章」显示「图 一-1」（所以汉字数字算数字）。首尾的分隔符去掉（「1.」→「1」）。
 * 罗马数字与字母编号（`A.1`）没算进来 —— 它们与「Chapter」这类文字分不开，宁可少留。
 *
 * 钉死办法：一份 docx，标题 1 分别挂「第1章」「第一章」「1.」「Chapter A」四种多级编号，
 * 每章里放一个 `{ STYLEREF 1 \s }` 和一个 `{ STYLEREF 1 \n \t }`，F9 后看显示什么。
 */
export const STYLEREF_NUMERIC_CHARS = /[0-9〇零一二三四五六七八九十百千壹贰叁肆伍陆柒捌玖拾佰仟.\-–—:]/u;

/**
 * 脚注分隔线的长度，twips。取 2 英寸（144pt）：英文版 Word 在 Letter 纸上那条线约占版心的三分之一，
 * 中文版 A4 上看着也是这个长度 —— 但都是**看**出来的，没有量过。
 *
 * 钉死办法：一份带两条脚注的 docx（A4 与 Letter 各一节）导 PDF，`truth.json` 的
 * `pages[].rules[]` 直接给出那条线的起止 x 与 y；顺带把下一条 `FOOTNOTE_SEPARATOR_LINE_Y` 一起钉死。
 *
 * 尾注开头那条短线也用它（与下一条）；尾注跨页时续页顶上那条按**通栏**画，同样没量过 ——
 * 同一份样本再加一段长过一页的尾注就能一起钉死。
 */
export const FOOTNOTE_SEPARATOR_WIDTH = 2880;

/**
 * 分隔线画在分隔线那一段（`footnotes.xml` 里 `separator` 那条的空段落）的什么高度，
 * 按那一段总高的比例。取 0.5（居中）。只影响画在哪，**不改任何一行的坐标** ——
 * 那一段占多高由它自己的段落属性决定，是已经标定过的行高规则。
 * 钉死办法同上。
 */
export const FOOTNOTE_SEPARATOR_LINE_Y = 0.5;

/**
 * 一条脚注放不下时，引用那一页**至少**要收它几行，才肯把它切开、剩下的续到下一页；
 * 连这几行都放不下，就把引用所在的那一行正文连同整条脚注挪到下一页（原来一律这么做）。
 * 取 1：「引用与脚注的第一行同页」是 Word 用户手册里的说法，但没有量过 ——
 * Word 也可能看孤行寡行（脚注文字样式继承 Normal，默认开着），那样该取 2。
 * 切口也不看脚注里各段的孤行寡行，同样没有真值。
 *
 * 钉死办法：一页正文里第 k 行引一条 6 行的脚注，k 从页底往上逐行挪（让本页剩下 0 / 1 / 2 / 3 行地方），
 * 导 PDF 看每一档是「整行挪走」还是「切开续排」、切开时本页留了几行；
 * 再把脚注的段落改成关掉孤行寡行比一遍。续页顶上那条通栏线的长度与位置用同一份样本的 `rules[]` 一起钉死。
 */
export const FOOTNOTE_SPLIT_MIN_LINES = 1;

/**
 * 方形环绕时，对象旁边的空当窄过几个字就不往里放字（单位 em，一个字宽取段落首字的字号）。
 *
 * 取 1：「至少放得下一个字」是能想到的最松的下限 —— 更松的话断行会在一格空当里硬塞一个字、
 * 溢出到对象上。Word 很可能有更大的下限（图贴着版心边时，剩下半个字宽的缝它肯定不放字），
 * 但到底是 1 个字、1 英寸还是按版心宽的比例，没有样本。
 *
 * 钉死办法：一张方形环绕的图，右边缘离版心右边 0.5 / 1 / 1.5 / 2 / 3 个字宽各一份（五段相同文字），
 * 看哪一份开始往图右边的缝里放字；换一个字号再做一遍，分得开「按字」还是「按绝对长度」。
 */
export const WRAP_MIN_SEGMENT_EM = 1;

/**
 * `wrapText="bothSides"`（规范默认，Word 界面里的「四周型」）两侧都放得下字时怎么排。
 *
 * Word 的做法是**一行分两段**：左边那段排满接着排右边那段，同一个 y 上两段文字。
 * 我们的一行只有一段（`LineLayout` 是一个 x 起点加一串片段，分页、命中测试、上下导航都按
 * 「一行一段」写的），所以退化成 `largest`：只走宽的那一侧，另一侧空着。
 * 图贴着版心一边放时（最常见的摆法）窄的那一侧本来就放不下字，两种排法一样；
 * 图居中时我们会多排出几行。遇到这种情形分页会记一条 `wrap-both-sides-approximated`。
 *
 * 这不是缺样本，是没实现：要改成两段，得让 `LineLayout` 支持同一个 y 上的第二段
 * （分页把两段当一行量高、索引按 x 分开两段）。
 */
export const WRAP_BOTH_SIDES_AS = 'largest';
