/**
 * 渲染层未标定的常数 —— 和 `@uw/layout` 的 `uncalibrated.ts` 同一条规矩：
 * 凡不是从 Word 真值反推出来的系数，一律关在这一个文件里，并写清「拿什么样本能钉死」。
 * 散落在画法代码里的魔法数字会被后人当成实测结论。
 *
 * 与布局那边的一个重要区别：**这里的数一个都不改坐标**。下划线画粗一点、上标抬高一点，
 * 断行点与基线 y 都不会动，所以 L2 / L3 的真值断言管不到它们 —— 这既是它们能拖到现在
 * 的原因，也是它们必须被单独记下来的原因（真值全绿 ≠ 画得对）。
 *
 * 钉死它们要的样本与已有的几份不同：现在的 `*.truth.json` 只抽**文字片段的 transform**，
 * 抽不到线条与装饰。要么给 `apps/fidelity` 的抽取器加一条「取 PDF 里的画线操作」，
 * 要么退一步做像素比对（把 Word 导出的 PDF 渲成位图，与本渲染器的 SVG 渲成的位图比）。
 * 前者精确、工作量大；后者便宜，但对不上时说不清是哪一条常数错了。
 */

/**
 * 下划线：基线往下多少（em）、画多粗（em）。
 *
 * 正确来源其实是字体的 `post.underlinePosition` / `underlineThickness`，但
 * `RawFontMetrics` 现在只收了 OS/2 与 hhea 两张表（行高与基线只用得着那两张），
 * 度量包里也就没有这两个数。补法很清楚：`readRawMetrics()` 多读一张 post 表、
 * 度量包格式加两个字段、重抽 17 款包 —— 那时这两条常数就可以删掉。
 */
export const UNDERLINE_OFFSET_EM = 0.09;
export const UNDERLINE_THICKNESS_EM = 0.06;

/**
 * 删除线：基线往**上**多少（em）。双删除线两条线的中心间距同样按 em 算。
 *
 * 同上，真正的来源是 OS/2 的 `yStrikeoutPosition` / `yStrikeoutSize`，度量包里没收。
 */
export const STRIKE_POSITION_EM = 0.26;
export const DOUBLE_STRIKE_GAP_EM = 0.09;

/**
 * 上下标的**升降量**，单位是该片段**已经缩过的**字号的 em
 * （字号系数 `VERT_ALIGN_SCALE = 2/3` 在 `@uw/layout` 那边，已经折进了 `fontSize`）。
 *
 * 布局侧的注释一直写着「升降量是另一回事」，那件事就是这两个数：它只改画的位置，
 * 不改任何宽度，所以到现在都没人被迫去量它。
 *
 * 钉死办法：一段 `x²` 与 `x₂`，导出 PDF 后看上标那个片段的 transform 的 **f 分量**
 * （基线 y）比正文基线高／低多少 pt，除以缩过的字号即可。现有的抽取器直接就能给出
 * 这个数 —— 这一条是本文件里最容易补上的。
 */
export const SUPERSCRIPT_RAISE_EM = 0.45;
export const SUBSCRIPT_DROP_EM = 0.14;

/**
 * 制表位前导符（目录里那排点）的画法。
 *
 * Word 画的是**一个个真实的字符**（`.` `-` `_`），我们画的是一条带 dash 图案的线 ——
 * 因为 `TabLeader` 只给了 x1 / x2，渲染层手上没有度量器，量不出一个点多宽、
 * 也就排不出「点与点之间对齐到哪」。视觉上非常接近，但**不是同一件事**：
 * Word 的点会与制表位右端对齐，我们的 dash 是从左端起算的。
 *
 * 钉死办法：一份带目录的样本，看真值里前导点片段的起始 x 是从左端还是右端排的；
 * 顺便量点的间距。真要做对，得让 `TabLeader` 带上字号并把度量器传进渲染层 ——
 * 那是个接口变动，等有样本证明差别看得见了再做。
 */
export const LEADER_DOT_PITCH_EM = 0.5;
export const LEADER_THICKNESS_EM = 0.055;
/** 前导符没有自己的字号，退到这个字号（12pt）—— 行内一个字都没有时才会走到 */
export const LEADER_FALLBACK_SIZE = 240;
