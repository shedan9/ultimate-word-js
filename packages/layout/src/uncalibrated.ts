/**
 * 未标定的常数 —— 全项目**唯一**允许出现「拍脑袋数字」的地方。
 *
 * 规矩（原则 1.5 的推论）：布局里任何一个不是从 Word 真值反推出来的系数，都必须放进这个文件，
 * 并写清楚「拿什么样本能把它钉死」。散落在各处的魔法数字会被后人当成实测结论，
 * 那比留一个洞危险得多 —— metrics.ts 里的 1.3 之所以可信，正因为它旁边贴着 13 个样本的误差表。
 *
 * 这里的值影响的是**宽度**，也就是断行位置 —— 行盒与基线已经标定完了（`@uw/fonts`
 * 的 `lineMetrics` / `baselineOffset`），现在挡着 L2（断行点与真值一致）的就是这几条。
 * 度量包进来之后 `layout/src/fixture.test.ts` 能逐行比真值了，18 行对上 8 行，
 * 差的那 10 行全指向下面的 `PUNCT_COMPRESS_RATIO`。
 * 末尾的 `BORDER_STYLE_RANK` 是个例外：它影响的是**画哪条线**，不改坐标，
 * 但同样是拍脑袋来的，所以一并关在这里。
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
 * gongwen-01 的真值**测不了它**：那份文档的中英之间本来就打了空格，
 * 真值里相邻中西片段的间隙精确为 0.000pt，也就是 Word 一点自动间距都没加。
 * 有空格时不加、没空格时加多少，是两个问题，这份样本只回答了前一个。
 *
 * 钉死办法：一行「中文English中文」，**中间不打空格**，量拉丁段的起始 x
 * 减去前一个汉字的右边缘。顺手把「打了空格时加不加」也写进同一份样本作对照。
 */
export const AUTO_SPACE_EM = 1 / 8;

/**
 * **塞不下时**临时挤标点的上限，单位是该标点自身宽度的比例。
 *
 * 注意这不是「相邻标点固定挤半个字」那条规则 —— 那一条已经实测钉死并搬去了
 * `break-class.ts` 的 `PUNCT_PAIR_COMPRESS_EM`（样本 `spike-punct-01`，26 段，误差 0.006 em）。
 * 留在这里的是另一件事：一行塞不下时，Word 会**额外**再挤一点行内的标点把字留住。
 *
 * 已经从真值反推出来的（`linebreak.ts` 的 `compress()` 按它实现）：
 * 压**整行**的标点而不只是正在溢出的那个字，且只压到刚好够。证据是 gongwen-01
 * 「为落实……通知如下。自」那一行 —— 26 个三号字自然宽 415.5pt、可用宽 410.25pt，
 * Word 挤出 6pt 把第 26 个字留在了行内，而那一行两个孤立标点合起来能挤 16pt。
 *
 * ⚠️ 没标定的是这个**上限**：实测只见过 0.19 em（那 6pt 分摊到两个标点），0.5 只是
 * 「不可能超过空半边」的物理上界。要钉死得造一行「差得多一点点、必须把标点挤到极限才留得住」
 * 的样本，看 Word 是挤到 0.5 还是宁可换行。
 *
 * ⚠️ 另有一条**已经量到但还没实现**的：悬挂标点的**墨留在版心内**，只有空半边吐出去。
 * gongwen-01 第 5 行的「，」左边缘落在版心线内 7.96pt、右边缘出界 8.05pt，正好一半一半；
 * 而我们现在把悬挂的标点整个不计入行宽。这是 L2 剩下 10 行对不上的直接原因（`fixture.test.ts`），
 * 修法是让悬挂项按半宽计入行宽 —— 它会同时改动两端对齐与行宽断言，所以单独做一步。
 */
export const PUNCT_COMPRESS_RATIO = 0.5;

/**
 * `w:*Chars`（1/100 字符）里「一个字符」的宽度 = 字号 × 这个系数。
 *
 * 取 1.0：中文字体的汉字推进宽度精确等于 1 em，公文的「首行缩进 2 字符」正是 2 个汉字宽。
 * 不去实际测一个汉字的宽度，是因为段落可能一个汉字都没有（纯英文段也能写 firstLineChars），
 * 那时按字号算才是 Word 的行为。
 */
export const CHAR_UNIT_EM = 1.0;

/** 把「em 的倍数」换成 twips。字号本身就是 twips，所以是一次乘法 */
export function em(fontSize: Twips, ratio: number): Twips {
  return fontSize * ratio;
}

/**
 * 边框冲突里「线宽一样时谁赢」的样式权重（大的赢）。
 *
 * 依据是 CSS 2.1 §17.6.2 的 collapsing borders：`double > solid > dashed > dotted`。
 * 拿它当 Word 的规则用是**类比**，不是实测 —— ECMA-376 只说 `w:tcBorders` 覆盖
 * `w:tblBorders`（§17.4.39），相邻两格谁赢一个字没提，而 Word 的表格边框行为
 * 整体上就是 collapsing 模型的变体。
 *
 * 认不出的 `w:val` 落到 `single` 那一档：实文件里的生僻线型（`dashDotStroked`
 * `thickThinMediumGap` …）都是实线的花样，退成点线会让它凭空输掉。
 *
 * 钉死办法：一张 2×2 的表，四条内部边分别让上下两格写不同的 `w:val`（宽度写成一样），
 * 导出 PDF 看画出来的是哪一种。一份样本能同时钉死这张表和下面的平局方向。
 */
const BORDER_STYLE_RANK: Record<string, number> = {
  // 多线型：视觉最重
  double: 3,
  triple: 3,
  doubleWave: 3,
  thickThinSmallGap: 3,
  thinThickSmallGap: 3,
  thickThinMediumGap: 3,
  thinThickMediumGap: 3,
  thickThinLargeGap: 3,
  thinThickLargeGap: 3,
  thinThickThinSmallGap: 3,
  thinThickThinMediumGap: 3,
  thinThickThinLargeGap: 3,
  // 实线
  single: 2,
  thick: 2,
  wave: 2,
  // 虚线族
  dashed: 1,
  dashSmallGap: 1,
  dotDash: 1,
  dotDotDash: 1,
  dashDotStroked: 1,
  // 点线最轻
  dotted: 0,
};

export function borderStyleRank(style: string): number {
  return BORDER_STYLE_RANK[style] ?? BORDER_STYLE_RANK.single ?? 2;
}
