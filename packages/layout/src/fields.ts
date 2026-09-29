/**
 * 域求值 —— 目前是 PAGE / NUMPAGES / SECTIONPAGES / PAGEREF，以及它们与分页之间的那个**循环**。
 *
 * 结构还原（界桩配对、指令解析）早在 `@uw/model` 的 fields.ts 做完了，这里只做「算成几」。
 * 分成两个包不是洁癖：配对只需要 run 序列，求值需要**页码**，而页码是分页的产物，
 * 所以它只能待在 layout 这一侧。
 *
 * ## 为什么必须迭代
 *
 * 页码依赖分页，域的文字宽度又反过来改断行、改分页：`{ PAGE }` 从 9 变成 10 多占半个字，
 * 一行少收一个字就可能多出一行，多出一行就可能多出一页，多出一页页码又变。
 * 所以这里是「排版 → 求值 → 再排版」的不动点迭代（开发计划 §2.4），不是一趟直算。
 *
 * 收敛判据是**入参**（那张 `FieldValues` 表）不再变：`layout = L(values)` 与
 * `values' = E(layout)` 都是确定性的，`E(L(values)) === values` 就说明手上这份 layout
 * 与它自己的页码自洽了。拿「页数不变」当判据是不够的 —— 页数一样、某个 PAGE 域从 3 变 4
 * 的情形完全可能（内容在页之间挪了位置）。
 *
 * ## 振荡检测（PAGEREF 进来之后才需要）
 *
 * 开发计划 §2.4 要求 A→B→A 的检测。只有 PAGE / NUMPAGES / SECTIONPAGES 时它**触发不了**：
 * 第一趟之后域文字只会变宽，分页规则只会把内容往后推，页码单调不减、回不了头。
 * **PAGEREF 打破了这一条**：目录里存的旧页码可以比实际的**大**（文档删短了、目录没更新），
 * 第一趟算出来的页码就会变窄，目录可能少一行、正文往前挪一页，页码又变 —— 两个解来回跳
 * 完全可能。所以每趟都比一遍「算出来的这份以前出现过没有」，出现过就是进了环，
 * 不必等撞上限，直接按计划说的「取页数较大者冻结」。
 * `MAX_FIELD_PASSES` 仍是最后一道保险丝。
 *
 * ## PAGEREF：书签落在第几页
 *
 * 目录（TOC）的页码不是 TOC 自己算的：Word 生成目录时给每条标题包一个 `_Toc…` 书签，
 * 条目里的页码是一个嵌套的 `PAGEREF _Toc… \h` 域。所以「目录页码正确」= 把 PAGEREF 算对，
 * 用不着重新生成整个目录（那是 TOC 求值，要大纲级别与样式，没做）。
 * 书签起点所在的段落从 `bookmarks`（`@uw/model` 的 `bookmarkTargets()`）查，页从布局查；
 * 显示的是**目标那一页**的显示页码、按**目标那一节**的页码格式 —— 前言用罗马数字的文档，
 * 目录里前言那几条就是 i、ii。`\p`（「见上方 / 见下方」）要一套本地化的措辞表，
 * 书签找不到时 Word 显示「错误！未定义书签。」—— 两种都**不求值**，照旧显示文件里存的结果，
 * 前者记 info、后者记 warn。
 *
 * ## 页眉页脚里的域走的是另一条路
 *
 * 同一个 `{ PAGE }` 在每一页显示的**不是同一串字**，所以一张全局的「run id → 文字」表
 * 按定义就装不下页眉里的域。它们改走 `LayoutDocumentOptions.headerFields`：那边存的是
 * **怎么算**（域类型 + 数字格式），算在**开页那一刻**做（`page.ts` 的 `headerFieldValues`）。
 * 直接的好处是页脚里的 PAGE **一趟就是准的** —— 页码在开页时已经定了，不必等下一趟迭代。
 * 还要迭代的只剩 NUMPAGES / SECTIONPAGES：「一共几页」得整份排完才知道。
 *
 * ## 三处容易搞反
 *
 * 1. **求值的结果不写回模型**，而是外挂一张「run id → 显示的文字」的表交给排版
 *    （`LayoutDocumentOptions.fieldValues`）。写回去意味着每一趟迭代克隆一棵树，
 *    而且模型里就有了两份真相：文件里存的旧值和我们算的新值，回写 docx 时不知道听谁的
 * 2. **没有 separate 的域不求值**。它在 Word 里就是什么都不显示（结果区都不存在），
 *    我们凭空往 begin 那个 run 上塞一串数字，等于替 Word 决定它该显示什么 ——
 *    没有样本支持的地方留洞，不猜（原则 1.5）。这种域会记一条 `field-no-result` 诊断
 * 3. **`NUMPAGES` 数的是物理页**，`PAGE` 用的是**显示页码**（`PageLayout.number`）。
 *    `w:pgNumType w:start` 会让某一节的页码重新起算，两者从那以后就对不上了
 */
import type { DiagnosticSink } from '@uw/core';
import type {
  BookmarkTargets,
  FieldInstruction,
  FieldRegion,
  NodeId,
  ResolvedBlock,
  ResolvedBody,
} from '@uw/model';
import { formatNumber, walkBlocks, walkParagraphs } from '@uw/model';
import type { HeaderFooterSource } from './header-footer.ts';
import type {
  DocumentLayout,
  HeaderFieldPlan,
  HeaderFieldSpec,
  LayoutDocumentOptions,
  PageLayout,
} from './page.ts';
import { layoutDocument } from './page.ts';
import type { BlockLayout } from './table.ts';
import type { LineLayout, ParagraphLayout } from './types.ts';
import { FIELD_CHINESE_NUM_FORMATS } from './uncalibrated.ts';

/** 求值结果：**run id → 这个 run 显示的文字**，盖掉它 content 里存着的旧值 */
export type FieldValues = ReadonlyMap<NodeId, string>;

/**
 * 能求值的域。其余一律**原样显示文件里存着的旧结果** —— Word 存盘时把上次算出来的文字
 * 写在 separate 与 end 之间，直接显示就是「打开即所见」，这也是本阶段之前不做求值
 * 也能正确渲染的原因。
 */
const EVALUABLE = new Set(['PAGE', 'NUMPAGES', 'SECTIONPAGES', 'PAGEREF']);

/**
 * 迭代上限。Word 自己 2–3 趟就收敛，5 是留够余量后的硬闸 ——
 * 它不是精度参数，是「别死循环」的保险丝。
 */
export const MAX_FIELD_PASSES = 5;

export interface LayoutDocumentWithFieldsOptions extends LayoutDocumentOptions {
  /** 最多排几趟，缺省 `MAX_FIELD_PASSES`。调它只有测试与调试用得上 */
  maxPasses?: number;
  /**
   * 书签名 → 起点所在段落（`@uw/model` 的 `bookmarkTargets(body)`）。PAGEREF 靠它找目标；
   * 缺席时 PAGEREF 一律不求值（显示文件里存的旧页码）
   */
  bookmarks?: BookmarkTargets;
}

export interface FieldLayoutResult {
  layout: DocumentLayout;
  /** 与 `layout` 自洽的那份求值结果。想复现这份布局，把它当 `fieldValues` 传回去即可 */
  values: FieldValues;
  /** 一共排了几趟。1 = 文档里没有可求值的域（或者文件里存的旧值已经全对） */
  passes: number;
  /** 撞上 `maxPasses` 时为 false —— 手上这份是「页数最多」的那一趟，不是收敛解 */
  converged: boolean;
}

/**
 * 排版 + 域求值，迭代到自洽。
 *
 * `fields` 来自 `loadDocument()` 的同名字段（`scanFields()` 的产物）。没有域、
 * 或者域全是不认识的类型时，这个函数与直接调 `layoutDocument()` 完全等价。
 */
export function layoutDocumentWithFields(
  body: ResolvedBody,
  fields: readonly FieldRegion[],
  opts: LayoutDocumentWithFieldsOptions,
): FieldLayoutResult {
  const { anchors, plan } = fieldAnchors(body, fields, opts.headerFooters, opts.bookmarks, opts.diagnostics);

  let values: FieldValues = opts.fieldValues ?? new Map();
  let totals: Totals = {};
  // 「一共几页」只有 NUMPAGES / SECTIONPAGES 用得上。全篇只有 PAGE 时不把它算进收敛判据，
  // 否则每份带页码的文档都要白排一趟（PAGE 在开页那一刻就是准的）
  const needsTotals = [...plan.fields.values()].some((f) => f.type !== 'PAGE' && f.type !== 'clear');

  // 诊断只在第一趟收：布局自己发的那几条（多栏、连续分节符改了版心）与域文字无关，
  // 每趟都发一遍只会让同一句话在诊断列表里出现三次
  let layout = layoutDocument(body, { ...opts, fieldValues: values, headerFields: { ...plan, ...totals } });
  if (anchors.length === 0 && !needsTotals) return { layout, values, passes: 1, converged: true };

  const { diagnostics: _quieted, ...quiet } = opts;
  const max = Math.max(1, opts.maxPasses ?? MAX_FIELD_PASSES);
  const tried: { values: FieldValues; totals: Totals; layout: DocumentLayout }[] = [
    { values, totals, layout },
  ];

  for (let passes = 1; ; passes++) {
    const next = evaluate(anchors, body, layout);
    const nextTotals: Totals = needsTotals
      ? { totalPages: layout.pages.length, sectionPages: countBySection(layout) }
      : {};
    if (sameValues(next, values) && sameTotals(nextTotals, totals)) {
      return { layout, values, passes, converged: true };
    }

    // 算出来的这份早先排过：再排只会绕回去（A→B→A），见文件头「振荡检测」
    // 最后一趟就是手上这份，与它相同已经在上面按收敛返回了，只比更早的
    const cycled = tried
      .slice(0, -1)
      .some((t) => sameValues(t.values, next) && sameTotals(t.totals, nextTotals));
    if (cycled || passes >= max) {
      opts.diagnostics?.warn(
        'field-not-converged',
        cycled
          ? `域求值在两个解之间来回跳（第 ${passes} 趟），冻结在页数最多的那一趟 —— 目录页码可能与 Word 差一页`
          : `域求值 ${max} 趟仍未收敛，冻结在页数最多的那一趟 —— 页码可能与 Word 差一页`,
      );
      // 计划 §2.4 的「取页数较大者冻结」：宁可多算一页也不要少算，
      // 少算的那一页会让最后一段内容整个消失，多算最多是末页留白
      const best = tried.reduce((a, b) => (b.layout.pages.length > a.layout.pages.length ? b : a));
      return { layout: best.layout, values: best.values, passes, converged: false };
    }

    values = next;
    totals = nextTotals;
    layout = layoutDocument(body, { ...quiet, fieldValues: values, headerFields: { ...plan, ...totals } });
    tried.push({ values, totals, layout });
  }
}

/** 迭代状态里「页数」的那一半，见 `needsTotals` */
interface Totals {
  totalPages?: number;
  sectionPages?: readonly number[];
}

function sameTotals(a: Totals, b: Totals): boolean {
  if (a.totalPages !== b.totalPages) return false;
  const x = a.sectionPages ?? [];
  const y = b.sectionPages ?? [];
  return x.length === y.length && x.every((n, i) => n === y[i]);
}

// ── 域 → 承载结果的那个 run ───────────────────────────────────────────────────

/** 一个待求值的域被压扁成的样子：算完往哪儿放、还要清掉谁 */
interface FieldAnchor {
  instruction: FieldInstruction;
  /** PAGEREF 的目标：书签起点所在的段落 */
  target?: NodeId;
  /** 结果文字写到这个 run 上 */
  runId: NodeId;
  /** 它所在的段落 —— 旧结果是空串时这个 run 排不出任何片段，只能靠段落定位 */
  paragraphId: NodeId;
  /** 结果区里其余的 run：Word 常把一个数字切成好几个 `w:t`，不清掉旧的会留在页面上 */
  clear: NodeId[];
}

/** 页眉页脚里所有 run 的 id —— 用来判断一个域落在正文还是落在页眉 */
function headerRunIds(source: HeaderFooterSource | undefined): Set<NodeId> {
  const out = new Set<NodeId>();
  if (source === undefined) return out;
  for (const content of Object.values(source)) {
    for (const b of walkBlocks(content.resolved as ResolvedBlock[])) {
      if (b.kind === 'paragraph') for (const run of b.runs) out.add(run.id);
    }
  }
  return out;
}

function fieldAnchors(
  body: ResolvedBody,
  fields: readonly FieldRegion[],
  headerFooters: HeaderFooterSource | undefined,
  bookmarks: BookmarkTargets | undefined,
  diagnostics?: DiagnosticSink,
): { anchors: FieldAnchor[]; plan: HeaderFieldPlan } {
  const runPara = new Map<NodeId, NodeId>();
  for (const p of walkParagraphs(body)) {
    for (const run of p.runs) runPara.set(run.id, p.id);
  }
  const inHeader = headerRunIds(headerFooters);

  const anchors: FieldAnchor[] = [];
  const planned = new Map<NodeId, HeaderFieldSpec>();
  const claimed = new Set<NodeId>();

  for (const region of fields) {
    const type = region.instruction.type;
    if (!EVALUABLE.has(type)) continue;

    const runId = region.resultRuns[0];
    if (runId === undefined) {
      diagnostics?.info(
        'field-no-result',
        `域 ${type} 没有结果区（缺 w:fldChar separate），Word 里它什么都不显示，这里也不求值`,
      );
      continue;
    }
    // 嵌套的可求值域（`{ PAGE { PAGE } }` 这种病态写法）：外层已经把这片 run 认领走了，
    // 内层再写一遍就成了两个域抢同一个 run。按文档顺序先到先得，后来的记诊断
    if (region.resultRuns.some((id) => claimed.has(id))) {
      diagnostics?.warn(
        'field-nested-eval',
        `域 ${type} 的结果区与外层域重叠，已跳过 —— 嵌套域的求值要先算内层，本阶段没做`,
      );
      continue;
    }
    const format = switchFormat(region.instruction);

    let target: NodeId | undefined;
    if (type === 'PAGEREF') {
      target = pageRefTarget(region.instruction, bookmarks, inHeader.has(runId), diagnostics);
      if (target === undefined) continue;
    }

    // 页眉页脚里的域走 plan 那条路：它每一页显示的不是同一串字，装不进一张全局表
    if (inHeader.has(runId)) {
      for (const id of region.resultRuns) claimed.add(id);
      planned.set(runId, {
        type: type as Exclude<HeaderFieldSpec['type'], 'clear'>,
        ...(format === undefined ? {} : { format }),
      });
      for (const id of region.resultRuns.slice(1)) planned.set(id, { type: 'clear' });
      continue;
    }

    const paragraphId = runPara.get(runId);
    // 结果 run 既不在正文也不在页眉：多半是脚注 / 批注这类还没解析的部件，没有页可谈
    if (paragraphId === undefined) continue;

    for (const id of region.resultRuns) claimed.add(id);
    anchors.push({
      instruction: region.instruction,
      ...(target === undefined ? {} : { target }),
      runId,
      paragraphId,
      clear: region.resultRuns.slice(1),
    });
  }
  return { anchors, plan: { fields: planned } };
}

/** PAGEREF 能不能求值、目标是哪一段。不能的记诊断并答 undefined（照旧显示存着的结果） */
function pageRefTarget(
  instr: FieldInstruction,
  bookmarks: BookmarkTargets | undefined,
  inHeader: boolean,
  diagnostics: DiagnosticSink | undefined,
): NodeId | undefined {
  const name = instr.args[0];
  if (bookmarks === undefined || name === undefined) return undefined;
  if (inHeader) {
    // 页眉里的 PAGEREF 每页算出来都一样，走 plan 那条路的「开页时算」反而说不清目标在哪一页 ——
    // 它在公文与报告里没见过，留洞
    diagnostics?.info('field-pageref-header', `页眉页脚里的 PAGEREF ${name} 不求值，显示文件里存的结果`);
    return undefined;
  }
  if (instr.switches.some((sw) => sw.name === 'p')) {
    diagnostics?.info(
      'field-pageref-relative',
      `PAGEREF ${name} \\p（见上方 / 见下方）不求值，显示文件里存的结果`,
    );
    return undefined;
  }
  const target = bookmarks.get(name);
  if (target === undefined) {
    diagnostics?.warn(
      'field-bookmark-missing',
      `PAGEREF 引用的书签「${name}」不存在（Word 更新域后会显示「错误！未定义书签。」），显示文件里存的结果`,
    );
  }
  return target;
}

// ── 一趟求值 ──────────────────────────────────────────────────────────────────

function evaluate(anchors: readonly FieldAnchor[], body: ResolvedBody, layout: DocumentLayout): FieldValues {
  const index = indexLayout(layout);
  const sectionPages = countBySection(layout);
  const out = new Map<NodeId, string>();

  for (const a of anchors) {
    // 旧结果是空串的域排不出任何片段，退到「它所在的段落排在第几页」。
    // 段落跨页时取它的**第一片**，与 Word 把域算在它自己那一行上略有出入 ——
    // 只有「域在一个跨页长段落的后半截」才看得出，公文里的页码都在独立的短段落里
    const pageIndex = index.runs.get(a.runId) ?? index.paragraphs.get(a.paragraphId);
    if (pageIndex === undefined) continue;
    // PAGEREF 显示的是**目标**那一页，不是域自己所在的那一页
    const shown = a.target === undefined ? pageIndex : index.paragraphs.get(a.target);
    if (shown === undefined) continue;
    const page = layout.pages[shown];
    if (page === undefined) continue;

    out.set(a.runId, fieldText(a.instruction, page, layout, sectionPages, body));
    for (const id of a.clear) out.set(id, '');
  }
  return out;
}

function fieldText(
  instr: FieldInstruction,
  page: PageLayout,
  layout: DocumentLayout,
  sectionPages: number[],
  body: ResolvedBody,
): string {
  const explicit = switchFormat(instr);
  switch (instr.type) {
    case 'PAGE':
    case 'PAGEREF': {
      // 没写 `\*` 时跟着**本节**的 `w:pgNumType w:fmt` —— 「前言用罗马数字、正文用阿拉伯
      // 数字」就是靠分节 + 那个属性实现的，忽略它整份前言的页码都会变成阿拉伯数字。
      // PAGEREF 的 page 是目标那一页，所以这里取的是**目标那一节**的格式
      const fmt = explicit ?? body.sections[page.sectionIndex]?.props.pageNumFormat ?? 'decimal';
      return formatNumber(page.number, fmt);
    }
    case 'NUMPAGES':
      // 数的是**物理页**，不是显示页码：pgNumStart 让页码从 5 起算时，
      // 「共几页」仍然是纸的张数
      return formatNumber(layout.pages.length, explicit ?? 'decimal');
    default:
      return formatNumber(sectionPages[page.sectionIndex] ?? 0, explicit ?? 'decimal');
  }
}

/**
 * `\*` 数字格式开关。
 *
 * 一条指令里可以有**两个** `\*`（Word 惯写 `PAGE \* roman \* MERGEFORMAT`），
 * 所以要挨个看，取第一个认得出的，而不是取第一个 `\*`。
 *
 * `roman` / `alphabetic` 的**大小写有意义**：Word 按开关值首字母的大小写决定输出大小写
 * （`\* ROMAN` → I、`\* roman` → i）。其余开关与大小写无关。
 *
 * 认不出的（`\* MERGEFORMAT` 只是「保留格式」不是数字格式，`\* ArabicDash` / `GB1`–`GB4`
 * 是「给数字套上破折号 / 括号」的包装写法，`\* CardText` 要一张英文数词表）一律退到
 * 十进制 —— 与 `formatNumber` 同一条规矩：**降级优于丢失**，页码错个样式是瑕疵，
 * 页码消失用户会以为文档坏了。
 */
function switchFormat(instr: FieldInstruction): string | undefined {
  for (const sw of instr.switches) {
    if (sw.name !== '*') continue;
    const raw = sw.value;
    if (raw === undefined || raw === '') continue;
    const key = raw.toLowerCase();
    const upper = raw.charAt(0) !== raw.charAt(0).toLowerCase();
    if (key === 'roman') return upper ? 'upperRoman' : 'lowerRoman';
    if (key === 'alphabetic') return upper ? 'upperLetter' : 'lowerLetter';
    if (key === 'arabic') return 'decimal';
    if (key === 'ordinal') return 'ordinal';
    const chinese = FIELD_CHINESE_NUM_FORMATS[key];
    if (chinese !== undefined) return chinese;
  }
  return undefined;
}

// ── 从布局里查「谁在第几页」 ──────────────────────────────────────────────────

interface LayoutIndex {
  /** run id → 它**第一次**出现在第几页（物理页序，0 起） */
  runs: Map<NodeId, number>;
  paragraphs: Map<NodeId, number>;
}

function indexLayout(layout: DocumentLayout): LayoutIndex {
  const index: LayoutIndex = { runs: new Map(), paragraphs: new Map() };
  layout.pages.forEach((page, i) => {
    for (const block of page.blocks) {
      if (block.kind === 'paragraph') {
        addParagraph(
          index,
          block.id,
          block.lines.map((l) => l.line),
          i,
        );
        continue;
      }
      // 表格：行是原子的（整行在同一页上），格内的块直接按这一页记
      for (const placed of block.rows) {
        for (const cell of placed.row.cells) addBlocks(index, cell.blocks, i);
      }
    }
  });
  return index;
}

function addBlocks(index: LayoutIndex, blocks: readonly BlockLayout[], page: number): void {
  for (const b of blocks) {
    if (b.kind === 'paragraph') addParagraphLayout(index, b.layout, page);
    else for (const row of b.layout.rows) for (const cell of row.cells) addBlocks(index, cell.blocks, page);
  }
}

function addParagraphLayout(index: LayoutIndex, p: ParagraphLayout, page: number): void {
  addParagraph(index, p.paragraphId, p.lines, page);
}

function addParagraph(
  index: LayoutIndex,
  paragraphId: NodeId,
  lines: readonly LineLayout[],
  page: number,
): void {
  // 一律**第一次**为准：跨页的段落在两页上各出现一次，域该算在它先出现的那一页
  if (!index.paragraphs.has(paragraphId)) index.paragraphs.set(paragraphId, page);
  for (const line of lines) {
    for (const f of line.fragments) if (!index.runs.has(f.runId)) index.runs.set(f.runId, page);
  }
}

/** 每一节各占几页。`filler` 那种为凑奇偶补出来的空页也算，它确实是一张纸 */
function countBySection(layout: DocumentLayout): number[] {
  const out: number[] = [];
  for (const page of layout.pages) out[page.sectionIndex] = (out[page.sectionIndex] ?? 0) + 1;
  return out;
}

function sameValues(a: FieldValues, b: FieldValues): boolean {
  if (a.size !== b.size) return false;
  for (const [k, v] of a) if (b.get(k) !== v) return false;
  return true;
}
