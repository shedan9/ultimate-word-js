/**
 * 行高与基线 —— 一行占多高，以及文字画在这个高度里的哪个位置。
 *
 * 两问都已由穿刺定死，实现分两边：**总量**在 `@uw/fonts` 的 `lineMetrics`（Phase 0，13 个样本），
 * **基线**在同一个包的 `baselineOffset`（基线穿刺，26 个样本）。这里做的是它们之上的四件事：
 * 按 `w:lineRule` 解释 `w:line`、多字体合成、按行网格吸附、把段落标记的度量补上。
 *
 * 顺序是这一层最容易搞反的地方，见 `applyLineRule` 的注释：网格吸附在**行距倍数之前**。
 */
import type { Twips } from '@uw/core';
import type { ComposeRule, LineMetrics, TextMeasurer } from '@uw/fonts';
import { baselineOffsetExact, bucketFont, composeLineBox } from '@uw/fonts';
import type { DocGrid, ResolvedParaProps } from '@uw/model';
import type { LayoutItem } from './types.ts';

export interface LineHeightContext {
  measurer: TextMeasurer;
  docGrid: DocGrid;
  defaultFont?: string;
  /** 内嵌对象的行盒规则。**标定用的接缝**，正常调用不要传，见 `OBJECT_RULES` */
  objectRules?: ObjectRules;
  /** 脚本与合成规则。同样是**标定用的接缝**，见 `SCRIPT_RULES` */
  scriptRules?: ScriptRules;
}

/**
 * 「这一行算东亚行还是拉丁行」与「几款字体怎么合成一个行盒」——
 * 与 `ObjectRules` 同理，留成接缝只为让 `spike:script` 把 12 种组合各跑一遍。
 */
export interface ScriptRules {
  /**
   * 行高走东亚规则（×1.3）还是拉丁规则（GDI 外部行距），按什么判：
   * - `font`：按**这一段文字实际用的那款字体**是不是东亚字体，**逐段**判（实测）
   * - `line`：按**这一行**有没有东亚字符，整行同一套（**旧实现**）
   */
  eastAsianBy: 'font' | 'line';
  /** 几款字体怎么合成一个行盒，见 `@uw/fonts` 的 `composeLineBox()` */
  compose: ComposeRule;
  /**
   * 哪些文字段参与行盒：
   * - `all`：全部（实测）
   * - `eastAsiaBucket`：这一行有东亚字符时，只有 **eastAsia 桶**的段参与，其余只当兜底下限（**旧实现**）
   */
  box: 'all' | 'eastAsiaBucket';
}

/**
 * 脚本与合成规则的实测值。样本 `spike-script-01`
 * （`pnpm --filter @uw/fidelity spike:script`）：11 页、每页四段同格式连排，
 * 每页换一种「`w:ascii` 槽 × `w:eastAsia` 槽」的配法，正文一律 36pt。
 * **8 种组合逐页跑，实现的这一组唯一满分 11/11**（第二名 8/11）。
 *
 * ① **`eastAsianBy: 'font'`** —— 走哪一套行高规则，看的是**实际画字的那款字体**，
 *    不是这一行有没有东亚**字符**，也不是 `w:eastAsia` 槽里装的是谁：
 *
 *    | 36pt 一整行纯 ASCII | ascii 槽 | eastAsia 槽 | Word 实测行高 | 那款字体的东亚 / 拉丁规则 |
 *    |---|---|---|---|---|
 *    | P1 | Times New Roman | 宋体 | 41.40 | 51.83 / **41.40** |
 *    | P2 | 宋体 | 宋体 | 46.72 | **46.80** / 41.06 |
 *    | P3 | 等线 | 等线 | 48.76 | **48.77** / 37.51 |
 *    | P4 | 微软雅黑 | 微软雅黑 | 61.77 | **61.77** / 47.51 |
 *    | P6 | Times New Roman | 等线 | 41.40 | 51.83 / **41.40** |
 *    | P7 | 微软雅黑 | 宋体 | 61.77 | **61.77** / 47.51 |
 *
 *    P2–P4、P7 把旧实现（按字符判 → 一律拉丁规则）打掉：一行里一个东亚字都没有，
 *    Word 照样按东亚规则算，**差 13–30%，而且每一行都差**。P1 / P6 把「按 eastAsia 槽判」
 *    打掉：那两页的 eastAsia 槽里是东亚字体，行高却是 Times 的拉丁值。
 *    `w:hint` 也不是答案 —— 同一页四段里 Word 自己写的 hint 有的带 `eastAsia` 有的不带，
 *    四段行高完全一样（差 0.12pt 以内，那是 PDF 的坐标噪声）。
 *
 * ② **`box: 'all'` + `compose: 'maxSides'`** —— 见 `composeLineBox()` 的证据表。
 *    P9/P10/P11 三页（等线画 ASCII、宋体 / 仿宋画汉字）实测 50.28pt，
 *    比**两款字体各自的行高都大**，旧实现的「只有东亚桶参与 + 取行高最大值」给 46.80，差 3.5pt。
 *
 * 这一条与基线穿刺的「东亚行的行盒只由东亚桶的字体决定」**不矛盾，是把它讲对了**：
 * spike-baseline-02 那页等线 72pt + Times 72pt，Times 的 winAscent 更大却没赢 ——
 * 原来解释成「拉丁 run 不参与」，实际是它参与了，但它作为**拉丁字体**走拉丁规则，
 * 核心盒上沿只有 67.22pt，输给等线走东亚规则的 69.57pt。两种说法在那一页上同解，
 * 而 P9 这种「两款东亚字体上下互不相让」的行只有新说法算得出来。
 */
export const SCRIPT_RULES: ScriptRules = {
  eastAsianBy: 'font',
  compose: 'maxSides',
  box: 'all',
};

/**
 * 内嵌对象（图片）在行盒里怎么摆。三条都是实测的，留成接缝只为了让
 * `apps/fidelity` 的 `spike:image` 能把八种组合各跑一遍，证明实现的这一组是唯一能复现 Word 的。
 */
export interface ObjectRules {
  /** 对象把行撑高之后，文字的下伸还留不留在基线以下 */
  keepDescent: boolean;
  /** `w:position` 对内嵌对象起不起作用 */
  raise: 'apply' | 'ignore';
  /** 对象在行盒里占的高度要不要按 1.5pt 量化（见 `objectBoxHeight`） */
  boxQuantum: 'round' | 'none';
  /**
   * 含内嵌对象的行怎么跟**行网格**与**倍数行距**打交道：
   * - `apart`：文字与对象**各算各的**再取大（实测，见 `lineHeight` 的两侧分算）
   * - `together`：两者合成一个自然行高，一起吸附、一起乘倍数（**旧实现**）
   * - `ignore`：含对象的行不吸网格
   */
  grid: 'apart' | 'together' | 'ignore';
}

/**
 * 内嵌对象行盒规则的实测值。样本 `spike-image-01`
 * （`pnpm --filter @uw/fidelity spike:image`）：仿宋 12pt 单倍行距，图高排成三条阶梯 ——
 * 粗的 4→60pt 十三档、细的 30→36pt 步长 0.5pt、微的 30.0→31.5pt 步长 0.1pt，
 * 另有 22pt 字号的两档对照与 `w:position` ±6pt 的两行。**44 个样本，最大偏差 0.140pt**。
 *
 * ① **盒底坐在基线上**：对象不留西文 descender 的那一截（`OBJECT_RULES` 之外的判据是
 *    ④ 的量化 —— 严格说坐在基线上的是**盒**，图在盒里靠上放）。
 *
 * ② **文字的下伸留着**（`keepDescent`）：图比文字高时，行高 = 盒高 + **文字自己的下伸**，
 *    不是盒高本身。仿宋 12pt 的下伸是 3.52pt：40pt 的图那一行，行高实测 43.95pt。
 *    22pt 字号的对照行给的是 6.41pt（= 22pt 仿宋的下伸），所以下伸跟着**字号**走，
 *    不是一个常数。原来的实现让对象把整行吃掉（行高 = 图高），每有一张图就少 3.5pt，
 *    这个错会一路累积到后面每一行的基线上。
 *
 * ③ **`w:position` 对图片起作用**（`raise`）：±6pt 的两行里，图整个跟着升降
 *    （实测 +5.95 / −6.09pt），且行盒跟着变：压低 6pt 的那一行下伸变成 6pt（实测 6.04），
 *    抬高 6pt 的那一行上伸变成 26pt（实测 26.0）。
 *
 * ④ **盒高按 1.5pt 四舍五入**（`boxQuantum`）：见下面的 `objectBoxHeight`。最早只取偶数 pt
 *    的粗阶梯里，这条表现为「图高 ≡ 4 (mod 6) 的那几档凭空多抬半磅」，看着像噪声；
 *    补一条 0.1pt 步长的微阶梯才看出是台阶（30.7pt → 30.77，30.8pt → 31.52，此后一路平到
 *    31.5pt）。**没有它，一页里每有一张图就可能偏 0.75pt，且往下累积** ——
 *    十三张图的阶梯样本上累计到了 1.5pt，早已越过 L3 判据。
 *
 * ⑤ **文字与对象各算各的再取大**（`grid`）：样本 `spike-image-03`（同一个 spike 脚本，
 *    每页 22 行 → 网格 31.8pt，图高 10→90pt 十二档 + 倍数行距七档，43 行）。两条：
 *    - 含图的行**参与网格吸附**：吸的是「盒高 + 文字下伸」，吸到网格行的整数倍，
 *      富余照旧上下均分（与纯文字行同一条规则）。28pt 的图（盒 28.5 + 下伸 3.52 = 32.02）
 *      吸成两个网格行 63.6pt，60pt 那一档（63.52，比两行只矮 0.12pt）仍是两行 ——
 *      边界落在 ceil 上，不是四舍五入
 *    - **倍数行距不乘在图撑起来的那一截上**：网格 31.8pt + 1.5 倍 + 40pt 的图，
 *      按「合成一个自然行高再乘」（旧实现）得 95.4pt，Word 给的是 63.6pt。
 *      实测的算法是两侧分算：文字侧 = 吸附 → 乘倍数，对象侧 = 「对象要的高 + 倍数在文字侧
 *      多留的那段空白」再吸附，取大者为行的推进量；基线在**赢的那一侧的行盒**里居中
 *      （对象侧的行盒**不含**那段空白 —— 关掉网格的两档里图底严丝合缝坐在基线上，
 *      多留的空白整个落在基线以下）
 */
export const OBJECT_RULES: ObjectRules = {
  keepDescent: true,
  raise: 'apply',
  boxQuantum: 'round',
  grid: 'apart',
};

/**
 * 对象在行盒里占的高度的量化刻度：**1.5pt**（= 30 twips = 96dpi 下的 2 个像素 = 1/48 英寸）。
 * 机理不明（Word 大概是在某个设备单位上取整），规律本身见 `objectBoxHeight`。
 */
const OBJECT_BOX_QUANTUM: Twips = 30;

/**
 * 对象在行盒里占的高度 = 图高**四舍五入**到 1.5pt 的整数倍，且**不小于图高本身**。
 *
 * 两半都是实测的：
 * - 四舍五入而不是向上取整：30.0–30.7pt 的图占的就是它自己那么高（余数不到半格不进位），
 *   30.8pt 起整个跳到 31.5pt 并一路平到 31.5pt —— 台阶边落在 30.75pt，正好是半格
 * - `max`：舍去的那一半会让图沉到基线以下，实测没有发生（8 / 20 / 32 / 35pt 这几档
 *   舍完的格子比图矮，量到的上伸仍是图高本身）
 *
 * 直接后果：图的**底边不总是严丝合缝坐在基线上**，盒底才是 —— 进位的那一档里图会浮在
 * 基线以上最多 0.75pt。所以 `LineObject.raise` 带的是「盒底减图底」加上 `w:position`。
 *
 * ⚠️ 恰好落在半格上（余数正好 0.5，如 30.75pt）舍向哪边没有样本 ——
 * 0.1pt 的阶梯跨过了 30.7 与 30.8，正好漏掉中间那一点。按 `Math.round` 的「.5 进位」处理。
 */
export function objectBoxHeight(height: Twips, rules: ObjectRules = OBJECT_RULES): Twips {
  if (rules.boxQuantum === 'none') return height;
  return Math.max(height, Math.round(height / OBJECT_BOX_QUANTUM) * OBJECT_BOX_QUANTUM);
}

export interface LineHeight {
  /** 行距规则与网格吸附之后的最终行高 */
  height: Twips;
  /** 行顶到基线，`0 <= baseline <= height` */
  baseline: Twips;
  /** 未经行距规则与网格调整的自然行高，回归比对时要能分辨「是规则的锅还是度量的锅」 */
  natural: Twips;
}

/**
 * 一行的行高与基线。
 *
 * 「走东亚规则还是拉丁规则」按**这一段文字用的那款字体**判、**逐段**判，不是按行里有没有
 * 东亚字符（实测，见 `SCRIPT_RULES` 第 ① 条）—— 一行「A2C6」若是用等线画的，
 * Word 照样按东亚规则给行高，差 30%。**行里的每一段都参与行盒**，各自的行盒逐项取 max
 * （第 ② 条）。
 *
 * 「东亚行的行盒只由东亚桶决定、拉丁 run 完全不参与」是这条规则的**旧说法**：
 * spike-baseline-02 那页等线 72pt + Times 72pt，Times 的 winAscent 更大却没赢 ——
 * 不是因为它没参与，而是因为它作为**拉丁字体**走拉丁规则、核心盒上沿只有 67.22pt，
 * 输给了等线的 69.57pt。两种说法在那一页上同解，`spike-script-01` 的 P9 才把它们分开。
 */
export function lineHeight(
  items: readonly LayoutItem[],
  range: { start: number; end: number },
  props: ResolvedParaProps,
  ctx: LineHeightContext,
): LineHeight {
  const rules = ctx.objectRules ?? OBJECT_RULES;
  const script = ctx.scriptRules ?? SCRIPT_RULES;
  const eastAsianLine = hasEastAsia(items, range);
  /** 定行盒的那些字体 */
  const box: LineMetrics[] = [];
  /** 不定行盒但也不能被切掉的那些，只用来兜底。`box: 'all'` 下恒为空 */
  const passenger: LineMetrics[] = [];
  const seen = new Set<string>();
  /** 对象要占的基线**以上** / **以下**各多少（`w:position` 已经算进去了） */
  let objectAbove = 0;
  let objectBelow = 0;
  /** 这一行有没有**撑得起高度**的内嵌对象。浮动对象在文字流里高宽都是 0，不算 */
  let hasObject = false;

  for (let i = range.start; i < range.end; i++) {
    const item = items[i] as LayoutItem;
    if (item.kind === 'object') {
      const raise = rules.raise === 'apply' ? (item.raise ?? 0) : 0;
      // 撑起行的是**盒**不是图：盒高按 1.5pt 四舍五入，图在盒里靠上放（见 `objectBoxHeight`）
      objectAbove = Math.max(objectAbove, objectBoxHeight(item.height, rules) + raise);
      objectBelow = Math.max(objectBelow, -raise);
      if (item.height > 0 || raise !== 0) hasObject = true;
      continue;
    }
    if (item.kind === 'break') continue;
    // 制表位也有字体（见 `TabItem.font`）：只有一个制表位的那一行行高全靠它
    const font = item.font === '' ? (ctx.defaultFont ?? '') : item.font;
    const key = `${font}|${item.fontSize}`;
    if (seen.has(key)) continue;
    seen.add(key);
    // 字体缺失时 `eastAsianFont` 答 undefined（「不知道」），退回按整行的字符判 ——
    // 谎报成拉丁字体会让一份缺字体的中文文档每一行都矮 30%
    const eastAsian =
      script.eastAsianBy === 'line' ? eastAsianLine : (ctx.measurer.eastAsianFont(font) ?? eastAsianLine);
    const m = ctx.measurer.lineMetrics(font, item.fontSize, { eastAsian });
    const decidesBox =
      script.box === 'all' || !eastAsianLine || (item.kind === 'char' && item.script === 'eastAsia');
    (decidesBox ? box : passenger).push(m);
  }

  if (box.length === 0) {
    // 空段落的行高全靠段落标记（¶）自己的字符属性，它不是摆设。
    // 「东亚行但一个东亚 item 都没有」不可能（eastAsianLine 就是这么判的），所以这里只会是空行。
    box.push(markMetrics(props, ctx));
  }

  // 内嵌对象坐在基线上（实测，见 `OBJECT_RULES`），所以它撑的是基线**以上**那一截；
  // 文字自己的下伸照旧留在基线以下 —— 行高 = 图高 + 文字下伸，不是图高。
  // 对象**不参与居中**：核心盒居中那条规则是给文字量的，把图片也居中会让它凭空浮起来。
  const floor = floorBox(passenger);
  const { natural: textNatural, above: textAbove } = composeLineBox(box, script.compose);
  const below = Math.max(rules.keepDescent ? textNatural - textAbove : 0, objectBelow);
  // 基线在**自然行高**下的位置：对象撑出来的高度是硬的，不许被「多出来的空间上下均分」摊掉，
  // 否则压低 6pt 的那一行会把基线又往下推半截（实测差 0.28pt）
  const above = Math.max(textAbove, objectAbove, floor.above);
  const natural = Math.max(textNatural, above + below, floor.height);
  const { height, box: baselineBox } = advance();
  return { height, baseline: baselineIn(height, baselineBox, props), natural };

  /**
   * 行的**推进量**与「基线在哪个盒子里居中」。没有内嵌对象时两者相等，就是老式子；
   * 有对象时分两侧算，这是 `spike-image-03` 量出来的（见 `OBJECT_RULES` 第 ⑤ 条）：
   *
   * - **文字侧**：吸附 → 乘倍数（顺序见 `applyGrid`，纯文字行的老规则）
   * - **对象侧**：对象要的高（盒高 + 文字下伸）**不乘倍数**，但倍数在文字侧多留的那段空白
   *   `extra` 会挂在它下面，然后一起吸附
   * - 取大者作推进量；基线在**赢的那一侧的行盒**里居中，而对象侧的行盒**不含** `extra`
   *
   * 最后那半句是关网格的两档逼出来的：1.5 倍 + 40pt 的图，图底严丝合缝坐在基线上
   * （实测 40.53 vs 盒高 40.5），多留的 7.8pt 整个落在基线以下 —— 把 `extra` 算进
   * 居中的盒子会让基线再往下沉 3.9pt。开着网格时这段空白被吸附吃掉了，所以只有关网格的
   * 样本能把它照出来。
   *
   * 平局归**文字侧**（`>` 而不是 `>=`）：2 倍行距 + 20pt 图那一档两侧都得 63.6pt，
   * 而基线实测在 40.06pt —— 那是文字侧的盒（63.6）算出来的，对象侧的盒（31.8）给的是 24.14。
   */
  function advance(): { height: Twips; box: Twips } {
    const snapped = applyGrid(natural, props, ctx.docGrid);
    // 固定值行距下 Word 就是**切**（见 `baselineIn`），对象不许把行撑开，也就没有两侧之说
    if (!hasObject || rules.grid === 'together' || props.spacing.lineRule === 'exact') {
      const h = applyLineRule(snapped, props);
      return { height: h, box: h };
    }
    const textSnapped = rules.grid === 'ignore' ? textNatural : applyGrid(textNatural, props, ctx.docGrid);
    const text = applyLineRule(textSnapped, props);
    const objectNatural = objectAbove + below;
    if (rules.grid === 'ignore') {
      return text > objectNatural
        ? { height: text, box: text }
        : { height: objectNatural, box: objectNatural };
    }
    // `extra` = 倍数**按文字的自然行高**多留出来的那段空白（单倍行距下就是 0，
    // 于是整个分支退化成老式子）。量的是**没吸附过的** `textNatural` —— 用吸附后的
    // 31.8pt 去乘，1.5 倍 + 20pt 图那一档会多算 8.1pt，正好把对象侧从 31.8 顶到 63.6，
    // 实测 Word 给的是文字侧的 47.7pt（`spike-image-03` 第 3 页第 2 行）。
    // 负数（倍数小于 1）夹到 0：对象要的高是硬下限，倍数不该把图压扁，这一半没有样本
    const extra = Math.max(0, applyLineRule(textNatural, props) - textNatural);
    const objectAdvance = applyGrid(objectNatural + extra, props, ctx.docGrid);
    if (objectAdvance <= text) return { height: text, box: text };
    return { height: objectAdvance, box: applyGrid(objectNatural, props, ctx.docGrid) };
  }

  /**
   * 固定值行距**不看字体**：基线就在行高的 80% 处（实测，见 `baselineOffsetExact`）。
   *
   * 连防切字的那两条下限（内嵌对象、拉丁 passenger）也一并不管 —— 用户写死了行高，
   * Word 的行为就是**切**（这正是「固定值」与「最小值」的区别）。在这里替他撑开，
   * 得到的页面会比 Word 少排几行，错得比切字更远。
   */
  function baselineIn(h: Twips, boxHeight: Twips, p: ResolvedParaProps): Twips {
    if (p.spacing.lineRule === 'exact') return baselineOffsetExact(h);
    // 行距倍数与网格吸附**多出来**的那部分才上下均分（基线穿刺的结论）。
    // 纯文字行里 `above + (h − natural) / 2` 与原来的 `composeBaseline(box, h)` 恒等 ——
    // `composeBaseline` 本身就是「核心盒在 h 里居中」，两种写法只在有对象时才分岔。
    // 居中用的是 `boxHeight` 而不是推进量 `h`：两者只在「对象侧赢 + 倍数行距」时不等，
    // 见 `advance()`
    return Math.min(h, above + (boxHeight - natural) / 2);
  }
}

/**
 * `w:lineRule` 三分支。`w:line` 的刻度取决于它，这也是 model 层故意不转这个单位的原因：
 * - `auto`：1/240 行，240 = 单倍
 * - `exact` / `atLeast`：twips
 *
 * 入参 `height` 是**已经吸附过网格**的高度 —— 顺序是实测的，见 `applyGrid`。
 *
 * ⚠️ `atLeast` 与网格的先后没有真值：这里把 `w:line` 当成吸附**之后**的下限
 * （`max(吸附后, line)`），另一种可能是先取下限再吸附。公文里 `atLeast` 罕见，
 * 一份「atLeast 20pt + 网格 31.8pt」的样本就能钉死，暂不为它设计。
 */
function applyLineRule(height: Twips, props: ResolvedParaProps): Twips {
  const { line, lineRule } = props.spacing;
  if (lineRule === 'exact') return line;
  if (lineRule === 'atLeast') return Math.max(height, line);
  return (height * line) / 240;
}

/**
 * 行网格吸附 —— 中文公文「每页 22 行」的实现，也是 Phase 0 穿刺踩过的坑：
 * 网格一开，基线被吸到 `linePitch` 的整数倍上，字体度量的差异整个被盖掉。
 *
 * 三条边界：
 * - `type="default"` 不吸（那就是「没有网格」）
 * - `w:lineRule="exact"`（固定值行距）不吸：用户既然写死了行高，网格不该再改它
 * - 段落关了 `w:snapToGrid` 不吸（实测：spike-baseline-03 第五段退回自然行高 20.76pt）
 *
 * **吸附在行距倍数之前**，这一条与直觉相反，是实测的（spike-baseline-03 末三段）：
 * 网格 31.8pt 下开 1.5 倍行距，仿宋 16pt 与宋体 12pt 的行高**都是 47.7pt**
 * （= 1.5 × 31.8），与字号无关。若是先乘倍数再吸附，两者分别是 31.2 与 23.4pt，
 * 都会吸到 31.8pt —— 差着半行，而且会随字号变。
 */
function applyGrid(height: Twips, props: ResolvedParaProps, grid: DocGrid): Twips {
  if (!props.snapToGrid) return height;
  if (grid.type === 'default' || grid.linePitch <= 0) return height;
  if (props.spacing.lineRule === 'exact') return height;
  const n = Math.max(1, Math.ceil(height / grid.linePitch - 1e-6));
  return n * grid.linePitch;
}

/**
 * 不定行盒的那些字体所需的最小 ascent —— 一条**防切字**的下限，不是实测规则。
 *
 * 实测样本里它永远不生效：中西混排且同字号时，东亚一侧算出来的基线（≥0.96 em）
 * 总是比拉丁一侧的核心盒上沿（≤0.94 em）低，所以加不加它，26 个样本的预测值一个都不变。
 * 加它是为了「12pt 汉字里嵌一个 72pt 英文单词」这种样本外的情形不至于让字叠在上一行 ——
 * 那时 Word 大概会把行撑高，但我们没有真值，宁可保守地保证不切字。
 *
 * 补一份「东亚小字号 + 拉丁大字号同行」的样本就能把这个 `Math.max` 换成真规则；
 * 在那之前它是**判断**，所以单独一个函数、单独一段注释，不许混进上面的实测公式里。
 */
function floorBox(parts: readonly LineMetrics[]): { above: Twips; height: Twips } {
  let above = 0;
  let height = 0;
  for (const p of parts) {
    if (p.coreAbove > above) above = p.coreAbove;
    // 下沿也要留出来，否则 g / y 的尾巴会被下一行盖掉
    if (p.coreAbove + p.descent > height) height = p.coreAbove + p.descent;
  }
  return { above, height };
}

function hasEastAsia(items: readonly LayoutItem[], range: { start: number; end: number }): boolean {
  for (let i = range.start; i < range.end; i++) {
    const item = items[i] as LayoutItem;
    if (item.kind === 'char' && item.script === 'eastAsia') return true;
  }
  return false;
}

/**
 * 空行的度量取段落标记的字符属性，走 **ascii 桶 + 拉丁规则**。
 *
 * 这是实测结论，且与直觉相反（spike-baseline-01 末两页）：一个只有段落标记的空段落，
 * 标记同时挂着 `w:eastAsia="宋体"` 与 `w:ascii="Times New Roman"`，Word 给的行高是
 * 13.82pt @12pt —— 那是 Times New Roman 的拉丁行高（1.1499 em），不是宋体的 15.6pt
 * （1.3 em）。空的黑体 22pt 段落同理给 25.34pt = Times 22pt，不是 28.6pt。
 *
 * 原因不神秘：段落标记本身不是东亚字符，逐字符分桶把它分到 ascii 桶，
 * 于是「这一行有没有东亚字」的答案是「没有」，`lineMetrics` 也就不该乘 1.3。
 * 换句话说这里不需要特例，需要的是**别自作聪明去优先东亚桶**。
 *
 * ⚠️ 未标定：`w:hint="eastAsia"` 会不会把答案翻过来。COM 不方便直接写 hint，
 * 要手改 XML 造样本；公文里空段落几乎都是这个配置，所以先按实测的来。
 */
function markMetrics(props: ResolvedParaProps, ctx: LineHeightContext): LineMetrics {
  const fonts = props.markRunProps.fonts;
  const ascii = bucketFont(fonts, 'ascii');
  const picked = ascii !== '' ? ascii : bucketFont(fonts, 'eastAsia');
  const font = picked === '' ? (ctx.defaultFont ?? '') : picked;
  // 走哪一套规则**跟着这款字体**，与正文的每一段同一条（`SCRIPT_RULES` 第 ① 条）。
  // 原来这里硬写 `eastAsian: false`，在 ascii 槽里装着 Times 的那两页上与新规则同解 ——
  // 「ascii 槽里装着东亚字体的空段落」这一格没有样本，但为它留一个与正文相反的
  // 特例，等于说「段落标记不算字」，那才是没有依据的一条。
  const script = ctx.scriptRules ?? SCRIPT_RULES;
  const eastAsian = script.eastAsianBy === 'line' ? false : (ctx.measurer.eastAsianFont(font) ?? false);
  return ctx.measurer.lineMetrics(font, props.markRunProps.size, { eastAsian });
}
