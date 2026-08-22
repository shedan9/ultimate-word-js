/**
 * 页眉页脚 —— 选哪一份、排成多高、摆在纸的什么位置，以及它怎么反过来挤正文。
 *
 * ## 它不是「另一个版心」
 *
 * 页眉页脚在 Word 里是**独立的框**：横向与版心同宽（左右页边距是共用的），纵向由
 * `w:pgMar` 的 `header` / `footer` 定位，内容多了就往版心那边长。三条几何规则：
 *
 * 1. 页眉框顶 = `w:header`（到**纸**顶的距离，不是到版心顶）
 * 2. 页脚框**底** = 纸高 − `w:footer` —— footer 量的是底边，与 header 量顶边**不对称**，
 *    按对称写会让页脚整体偏移一个「页脚高度」
 * 3. 页眉长过 `w:top` 时**把正文往下顶**（版心顶 = max(`w:top`, 页眉底)），页脚同理
 *    往上顶版心底。这是「页边距是最小值不是固定值」的直接后果
 *
 * 三条都由 `spike-header-01/02` 实测（见 `HEADER_RULES` 的证据表）。
 *
 * ## 为什么每一页各排一遍
 *
 * 同一份页眉在每一页画的**不一定是同一串字**：`{ PAGE }` 每页都不同。所以排版结果
 * 按「内容 + 这一页的域文字」做键缓存，纯静态的页眉（绝大多数）全文档只排一次，
 * 带页码的那种每页排一次。缓存命中时几页共用同一份数据对象 —— 只读，且
 * `structuredClone` 会保住这种共享，不会把一份复制成两百份。
 *
 * ## 还没有真值的四问（几何有，**选择**没有）
 *
 * 几何那三条由 `spike-header-01/02` 钉死（见 `HEADER_RULES`），`spike-header-03` 顺手证了
 * 「首页 → 偶 → 奇 → 偶 → 奇」这一串落位。下面四问那份样本覆盖不到 ——
 * 代码里**按规范实现了，但没有样本**，与 `cascade-table.ts` 里隔行带的序号同一个性质：
 *
 * ① **奇偶看的是显示页码还是物理页序**。03 的页码从 1 起算，两者恰好重合，分不开。
 *    钉死办法：同一份样本加 `w:pgNumType w:start="2"`，看第一张纸用奇数还是偶数那一份
 * ② **`w:titlePg` 开着却没定义 first 时首页是不是空的**。按 §17.10.6 实现成「空」——
 *    退回 default 会让每份带封面的公文首页多出一行页眉。钉死办法：把 03 的 first 那一份删掉
 * ③ **跨节沿用**：本节没写某一类引用时沿用前一节的（§17.10.5，界面上的「链接到上一节」）。
 *    钉死办法：两节的 docx，第二节只写 footer 不写 header，看第二节的页上有没有页眉
 * ④ **`evenPage` / `oddPage` 补出来的空页算不算「本节首页」**。现在算 —— 它确实是这一节的
 *    第一张纸。钉死办法：`evenPage` + `w:titlePg` 的两节样本，看那张空页上画的是哪一份
 *
 * 四问都**不改坐标**，只改「画的是哪一份」，所以优先级低于表格与图片；
 * 但公文封面几乎必用 ①②，真拿到多页公文语料时该一起做掉。
 */
import type { Twips } from '@uw/core';
import type { TextMeasurer } from '@uw/fonts';
import type { DocumentSettings, HeaderFooterRef, NodeId, ResolvedBlock, SectionProps } from '@uw/model';
import type { PageGeometry, PlacedBlock, PlacedLine } from './page.ts';
import { layoutParagraph } from './paragraph.ts';
import { layoutTable } from './table.ts';

/** 页眉 / 页脚的内容来源：关系 id → 级联完的块。形状与 `LoadedDocument.headerFooters` 对得上 */
export type HeaderFooterSource = Readonly<Record<string, { resolved: readonly ResolvedBlock[] }>>;

/** 一页上画出来的页眉或页脚。`x/y` 相对**纸**左上角，块的 y 相对这个框的顶 */
export interface PlacedHeaderFooter {
  kind: 'header' | 'footer';
  /** 引用它的关系 id —— 想回查内容 / 回写部件时靠它 */
  relId: string;
  x: Twips;
  y: Twips;
  width: Twips;
  height: Twips;
  blocks: PlacedBlock[];
}

/**
 * 这一页该用哪一份页眉 / 页脚。
 *
 * 两条容易搞反的：
 *
 * ① **`w:titlePg` 开着却没定义 first 时，首页是空的**，不是退回 default（§17.10.6）。
 *    「不同的首页」在 Word 界面上的语义就是「首页那一份单独给」，没给就是没有；
 *    退回 default 会让每份带封面的公文首页多出一行页眉
 * ② **缺某一类就沿用前一节的**（§17.10.5），所以要往回找而不是就地放弃 ——
 *    Word 的「链接到上一节」在文件里正是「本节干脆不写这一类引用」
 *
 * 奇偶用的是**显示页码**（`PageLayout.number`）而不是物理页序：`w:pgNumType w:start`
 * 让页码从 2 起算时，第一张纸要用偶数页那一份。没有真值，见文件尾的未标定清单。
 */
export function pickHeaderFooter(
  sections: readonly SectionProps[],
  sectionIndex: number,
  kind: 'header' | 'footer',
  firstInSection: boolean,
  pageNumber: number,
  settings: DocumentSettings,
): HeaderFooterRef | undefined {
  const own = sections[sectionIndex];
  if (own === undefined) return undefined;

  const wanted: HeaderFooterRef['type'] =
    own.titlePage && firstInSection
      ? 'first'
      : settings.evenAndOddHeaders && pageNumber % 2 === 0
        ? 'even'
        : 'default';

  for (let i = sectionIndex; i >= 0; i--) {
    const props = sections[i];
    if (props === undefined) continue;
    const refs = kind === 'header' ? props.headers : props.footers;
    const hit = refs.find((r) => r.type === wanted);
    if (hit !== undefined) return hit;
    // 只有「这一节压根没写这一类」才继续往回找。写了 default 不代表能顶替 first：
    // 那正是 ① 说的那条，所以这里比的是 wanted 而不是「有没有任何一条」
  }
  return undefined;
}

export interface StackOptions {
  measurer: TextMeasurer;
  settings: DocumentSettings;
  docGrid: SectionProps['docGrid'];
  contentWidth: Twips;
  defaultFont?: string;
  fieldValues?: ReadonlyMap<NodeId, string>;
}

/** 摞出来的结果：块 + 总高。页眉不分页，所以这里没有任何「放不下」的分支 */
export interface StackResult {
  blocks: PlacedBlock[];
  height: Twips;
}

/**
 * 把一列块从 y=0 往下摞。
 *
 * 与分页那条路的差别只有一个：**不会换页**，所以段前段后间距一律照加 ——
 * 「段前间距落在页首不算」那条规则是分页规则（`PAGINATION_RULES` ②），
 * 页眉里的第一段并不是「被分页顶到页首的」，它本来就在那儿。
 */
export function stackBlocks(blocks: readonly ResolvedBlock[], opts: StackOptions): StackResult {
  const shared = {
    measurer: opts.measurer,
    settings: opts.settings,
    docGrid: opts.docGrid,
    ...(opts.defaultFont === undefined ? {} : { defaultFont: opts.defaultFont }),
    ...(opts.fieldValues === undefined ? {} : { fieldValues: opts.fieldValues }),
  };

  const out: PlacedBlock[] = [];
  let y: Twips = 0;

  for (const b of blocks) {
    if (b.kind === 'paragraph') {
      const layout = layoutParagraph(b, { ...shared, contentWidth: opts.contentWidth });
      y += layout.spaceBefore;
      const top = y;
      const lines: PlacedLine[] = [];
      layout.lines.forEach((line, index) => {
        lines.push({ index, y, line });
        y += line.height;
      });
      out.push({ kind: 'paragraph', id: b.id, y: top, lines, first: true, last: true });
      y += layout.spaceAfter;
      continue;
    }

    const layout = layoutTable(b, { ...shared, availWidth: opts.contentWidth });
    const top = y;
    const rows = layout.rows.map((row, index) => {
      const placed = { index, y, height: row.height, row };
      y += row.height;
      return placed;
    });
    out.push({
      kind: 'table',
      id: b.id,
      x: layout.x,
      y: top,
      width: layout.width,
      columns: [...layout.columns],
      rows,
      first: true,
      last: true,
    });
  }

  return { blocks: out, height: y };
}

// ── 几何：页眉页脚怎么反过来挤版心 ────────────────────────────────────────────

/**
 * 页眉页脚定位与「顶开版心」的实测值。样本 `spike-header-01`（矮）与 `spike-header-02`（高），
 * 两份除页眉页脚的行数外逐字相同：120×120mm 的纸、四边 20mm、页眉页脚距各 10mm、
 * 固定行距 20pt 仿宋 12pt，于是一页恰好 11 行、行高恰好 20pt，几何里不掺任何待标定的度量。
 *
 * ① **页眉框顶 = `w:header`（到纸顶）**：01 的页眉基线 44.420pt，
 *    预测 28.346（10mm）+ 0.8 × 20（固定行距的基线，`baselineOffsetExact`）= 44.35pt，差 0.07pt。
 *
 * ② **页脚量的是框底：框底 = 纸高 − `w:footer`**，与页眉量顶边**不对称**。
 *    01 的页脚基线 307.896pt；按「量底」预测 340.157 − 28.346 − 20 + 16 = 307.81pt（差 0.09pt），
 *    按「量顶」预测 327.81pt —— **差 20pt**，这一条分得干干净净。
 *    02 的三行页脚基线 267.936 / 287.976 / 307.896：末行贴着同一条底线，
 *    整块是**往上长**的，再次印证量的是底边。
 *
 * ③ **页边距是最小值，不是固定值**：版心顶 = max(`w:top`, 页眉底)、版心底 = min(纸高 − `w:bottom`, 页脚顶)。
 *    两个方向在两份样本里各被证了一次 ——
 *    · 01（页眉底 48.35 < 上边距 56.69）：正文首行基线 72.740，与没有页眉的 `spike-page-01` 一模一样；
 *      一页 11 行（页脚顶 311.81 比下边距 283.46 低，min 取的是下边距）
 *    · 02（页眉底 88.35 > 上边距 56.69）：正文首行基线 **104.450**，预测 88.35 + 16 = 104.35，差 0.10pt；
 *      一页 **8 行** —— 若页脚不参与，版心底会是下边距 283.46，第 9 行落在 268.35 处**放得下**，
 *      Word 却换了页，而页脚顶 251.81 正好挡在第 9 行前面
 *
 * 四种「挤不挤」× 两种「页脚量哪一边」共 8 组候选里，只有 `both` + `bottom` 能同时解释
 * 这三份样本（`spike:header` 把 8 组逐页跑了一遍）：不挤 → 01 对 02 错；只挤顶 → 02 的行数
 * 错（会是 9 行）；只挤底 → 02 的首行基线差 16pt；页脚量顶边 → 三份的页脚全差一个页脚高度。
 */
export interface HeaderRules {
  /** `w:pgMar/@w:footer` 量到页脚的哪一边 */
  footerAnchor: 'bottom' | 'top';
  /** 页眉页脚长过页边距时顶不顶版心：两头都顶 / 都不顶 / 只顶上 / 只顶下 */
  squeeze: 'both' | 'none' | 'top' | 'bottom';
}

export const HEADER_RULES: HeaderRules = {
  footerAnchor: 'bottom',
  squeeze: 'both',
};

/**
 * 把页眉页脚的高度算进版心。
 *
 * `base` 是这一节的**纸面**几何（`pageGeometry()` 的产物，只减了页边距与装订线），
 * 返回的是**这一页**的几何 —— 同一节里各页的版心可以不一样高，因为首页与偶数页
 * 用的页眉可能不一样长。
 */
export function contentWithHeaderFooter(
  base: PageGeometry,
  margin: SectionProps['margin'],
  headerHeight: Twips,
  footerHeight: Twips,
  rules: HeaderRules = HEADER_RULES,
): PageGeometry {
  const squeezeTop = rules.squeeze === 'both' || rules.squeeze === 'top';
  const squeezeBottom = rules.squeeze === 'both' || rules.squeeze === 'bottom';
  const top = squeezeTop ? Math.max(base.content.y, margin.header + headerHeight) : base.content.y;
  const paperBottom = base.content.y + base.content.height;
  const bottom = squeezeBottom
    ? Math.min(paperBottom, footerTop(base, margin, footerHeight, rules))
    : paperBottom;
  return {
    width: base.width,
    height: base.height,
    content: {
      x: base.content.x,
      y: top,
      width: base.content.width,
      // 页眉页脚长到把版心吃光时钳到 0：负高度会让 `fitLines` 一行都放不下，
      // 于是每一页硬塞一行、无限翻页
      height: Math.max(0, bottom - top),
    },
  };
}

function footerTop(
  base: PageGeometry,
  margin: SectionProps['margin'],
  height: Twips,
  rules: HeaderRules,
): Twips {
  const bottomOfFooter = base.height - margin.footer;
  return rules.footerAnchor === 'bottom' ? bottomOfFooter - height : bottomOfFooter;
}

/** 页眉 / 页脚框在纸上的位置。页眉从 `w:header` 往下长，页脚从纸底往上长（见 ②） */
export function frameOf(
  kind: 'header' | 'footer',
  base: PageGeometry,
  margin: SectionProps['margin'],
  height: Twips,
  rules: HeaderRules = HEADER_RULES,
): { x: Twips; y: Twips; width: Twips; height: Twips } {
  const y = kind === 'header' ? margin.header : footerTop(base, margin, height, rules);
  return { x: base.content.x, y, width: base.content.width, height };
}
