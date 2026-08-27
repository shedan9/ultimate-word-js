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
 * `line-height.ts` 的 `OBJECT_RULES` 与 `page.ts` 的 `FLOAT_ORIGIN_RULES`。`layout/src/fixture.test.ts` 现在 18 行对上 16 行，
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
 * 中西文自动间距的宽度（`w:autoSpaceDE` / `w:autoSpaceDN`），单位是**东亚一侧字号的 em**。
 *
 * 1/8 em 是开发计划 §2.2 记下的值，但没有自己的真值样本 —— 它与行长直接相关，
 * 中英混排一行差几个百分点就会换断行点。
 *
 * gongwen-01 的真值**测不了这个数**：那份文档的汉字与西文之间本来就打了空格，
 * 真值里相邻中西片段的间隙精确为 0.000pt，也就是 Word 一点自动间距都没加。
 * （那个空格本身量到 7.95–8.04pt = 仿宋的 0.5 em，不是 Times 的 0.25 em ——
 * 见 `items.ts` 的 `applySpaceFont`。间隙为 0 这条结论正是拿它对出来的。）
 *
 * 它**回答了另一个问题**：`（ascii`、`cs）`、`（autoSpaceDE` 三处是全角标点直接挨着西文、
 * 中间没有空格，实测间隙都在 0.05pt 以内 —— **全角标点旁边一点都不加**。
 * 这一条已经落到 `applyAutoSpace` 里了，也是真值第 13 行能对上的原因之一。
 * 剩下没标定的就是「**汉字**直接挨着西文、中间没空格」时加多少。
 *
 * 钉死办法：一行「中文English中文」，**中间不打空格**，量拉丁段的起始 x
 * 减去前一个汉字的右边缘。顺手把「打了空格时加不加」也写进同一份样本作对照。
 */
export const AUTO_SPACE_EM = 1 / 8;

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
