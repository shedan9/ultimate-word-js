/**
 * 目录（TOC 域）重新生成 —— Word 里「更新目录 → 更新整个目录」那一下。
 *
 * 目录**页码**早就不用它了：条目里的页码是嵌套的 `PAGEREF _Toc… \h`，布局那侧按书签所在页求值
 * （`@uw/layout` 的 fields.ts）。这里管的是页码之外的那一半 —— **标题增删、改字之后条目跟着变**，
 * 那要把 TOC 的结果区整个换掉，是**改模型**，不是排版时外挂一张表能做的：条目多一条就多一个段落，
 * 段落要有 id、要能被点、要能回写，外挂表装不下。
 *
 * 所以分两半：这里是**纯函数**，从级联完的树上挑标题、拼出新条目的段落草稿（`TocPlan`）；
 * 落到树上的是事务的 `replaceFieldResult` / `addBookmark`，跟着撤销走、回写成 docx。
 * 事务看不到样式表与大纲级别（它只吃直接格式那棵树），这一步必须在事务外面算好。
 *
 * **不自动刷新**：Word 打开文档、编辑标题都不会动目录，要用户点「更新目录」（F9）。
 * 自动重排意味着每打一个字都往撤销栈里塞一次目录改写，而且与 Word 打开同一份文件看到的不一样。
 *
 * 条目的样子照 Word 2016 中文版生成的抄（`\h` 时整条包一层 `w:hyperlink w:anchor`）：
 * `[编号 分隔] 标题文字 <tab> { PAGEREF _Toc… \h }`，段落样式 `toc N`、一个版心宽处的右对齐点前导制表位。
 * 缺了 `toc N` 样式就按中文版默认模板补一份（`tocStyleDefinition`）。标题没有 `_Toc` 书签的
 * 补一个（名字接着文档里最大的那个编号往下数），已有的沿用 —— 目录跳转与页码都靠它。
 *
 * 几处**没有 Word 样本**、照规范与观察写的，钉死办法都是「Word 里更新一次目录、看 document.xml」：
 * 1. `\o "1-3"` 按**级联后的大纲级别**认（样式里写了 `w:outlineLvl` 的自定义样式也算），与 `\u` 同解。
 *    规范说 `\o` 认「内建标题样式」，但中文公文大量用自定义样式 + 大纲级别，Word 实际照收
 * 2. 编号与标题之间的分隔跟着编号自己的 `w:suff`（制表位 / 空格 / 无）
 * 3. 标题里的制表位与换行写成空格（`\w` / `\x` 保留），隐藏文字不进条目
 *
 * 不认的开关整个目录不更新（显示原样）：`\c` / `\a`（图表目录，按 SEQ 收）、`\f` / `\l`（TC 域）、
 * `\b`（只收某书签范围内的 —— 书签只记了起点段落）、`\s` / `\d`（页码前面带章节号）。
 */
import type { FieldInstruction, FieldRegion } from './fields.ts';
import { fieldSwitch } from './fields.ts';
import type { Body, NodeId, ResolvedBody, ResolvedParagraph, RunContent } from './nodes.ts';
import { walkParagraphs } from './nodes.ts';
import type { ParaProps, RunProps, TabStop } from './props.ts';
import type { StyleDefinition } from './styles.ts';

/** 事务新造段落用的草稿：没有 id（事务给），其余与 `Paragraph` 同形 */
export interface ParagraphDraft {
  props: ParaProps;
  runs: RunDraft[];
}

export interface RunDraft {
  props?: RunProps;
  content: RunContent[];
  /** 目录条目的超链接（`\h`）：跳到标题的书签 */
  hyperlink?: { anchor: string };
}

export interface TocPlan {
  /** TOC 域 begin 界桩所在的 run，交给 `tx.replaceFieldResult` */
  field: NodeId;
  paragraphs: ParagraphDraft[];
  /** 要补的书签（标题上还没有 `_Toc…` 的），先于 `replaceFieldResult` 加上 */
  bookmarks: { paragraphId: NodeId; name: string }[];
  /** 要补的 `toc N` 样式定义（文档里没有的），交给 `tx.addStyle` */
  styles: StyleDefinition[];
  /** 条目数；0 时 `paragraphs` 是那一段「未找到目录项。」 */
  entries: number;
}

export type TocPlanResult = { ok: true; plan: TocPlan } | { ok: false; reason: string };

export interface TocPlanOptions {
  /** 段落样式 id → `w:name`（`paragraphStyleNames()`）：认 `toc N` 与 `\t` 里的样式名 */
  styleNames: ReadonlyMap<string, string>;
  /** 默认段落样式 id，补 `toc N` 时 basedOn / next 指它；空串 = 不写 */
  normalStyleId: string;
  /** 标题段落当前显示的页码（按目标那一节的页码格式）。缺席写空串，排版时 PAGEREF 照样算出来 */
  pageText?: (paragraphId: NodeId) => string | undefined;
  /**
   * 同一个事务里先前几个目录已经分给标题的新书签（段落 id → 名字），当作已有的 ——
   * 一份文档两个目录时，两边各自从同一个编号往下数就撞名了
   */
  assigned?: ReadonlyMap<NodeId, string>;
}

/** Word 找不到目录项时结果区写的那一句（中文版原文） */
export const TOC_EMPTY_TEXT = '未找到目录项。';

const UNSUPPORTED: readonly [string, string][] = [
  ['c', '图表目录（\\c）'],
  ['a', '图表目录（\\a）'],
  ['f', 'TC 域目录（\\f）'],
  ['l', 'TC 域级别（\\l）'],
  ['b', '书签范围（\\b）'],
  ['s', '带章节号的页码（\\s）'],
  ['d', '章节号分隔符（\\d）'],
];

/** 中文版默认模板的 toc N：左缩进每级两字（420 twips），uiPriority 39 */
const TOC_INDENT_STEP = 420;

/**
 * 补一份 `toc N` 样式定义（1–9）。`taken` 答 id 是否已被占用，撞了加序号 ——
 * 与 `builtinStyleDefinition` 同一套规矩。
 */
export function tocStyleDefinition(
  level: number,
  taken: (id: string) => boolean,
  normalId: string,
): StyleDefinition {
  let id = `TOC${level}`;
  for (let n = 1; taken(id); n++) id = `TOC${level}_${n}`;
  const indent =
    level > 1 ? { left: TOC_INDENT_STEP * (level - 1), leftChars: 200 * (level - 1) } : undefined;
  return {
    id,
    name: `toc ${level}`,
    ...(normalId === '' ? {} : { basedOn: normalId, next: normalId }),
    paraProps: indent === undefined ? {} : { indent },
    runProps: {},
    uiPriority: 39,
  };
}

/** 文档里所有 TOC 域（按 begin 的先后），只收正文里有 begin / separate / end 的 */
export function tocFields(fields: readonly FieldRegion[]): FieldRegion[] {
  return fields.filter(
    (f) =>
      f.kind === 'complex' &&
      f.instruction.type === 'TOC' &&
      f.begin !== undefined &&
      f.separate !== undefined &&
      f.end !== undefined,
  );
}

/**
 * 算出一个 TOC 域的新结果区。`fields` 是整份正文的域（排除别的目录的结果区用），
 * `region` 是其中要更新的那一个。
 */
export function planTableOfContents(
  resolved: ResolvedBody,
  body: Body,
  fields: readonly FieldRegion[],
  region: FieldRegion,
  opts: TocPlanOptions,
): TocPlanResult {
  const instr = region.instruction;
  if (instr.type !== 'TOC' || region.begin === undefined || region.separate === undefined || !region.end)
    return { ok: false, reason: '不是完整的 TOC 域（缺 separate / end）' };
  for (const [name, what] of UNSUPPORTED) {
    if (fieldSwitch(instr, name) !== undefined) return { ok: false, reason: `不支持${what}` };
  }

  const select = levelSelector(instr, opts.styleNames);
  const inToc = tocResultParagraphs(resolved, fields);
  const bookmarkNames = new Set<string>();
  const onParagraph = new Map<NodeId, readonly string[]>();
  for (const p of walkParagraphs(body)) {
    if (p.bookmarks === undefined) continue;
    onParagraph.set(p.id, p.bookmarks);
    for (const b of p.bookmarks) bookmarkNames.add(b);
  }
  for (const [id, name] of opts.assigned ?? []) {
    onParagraph.set(id, [...(onParagraph.get(id) ?? []), name]);
    bookmarkNames.add(name);
  }
  let serial = nextTocSerial(bookmarkNames);

  const keepTabs = fieldSwitch(instr, 'w') !== undefined;
  const keepBreaks = fieldSwitch(instr, 'x') !== undefined;
  const noPages = levelRange(fieldSwitch(instr, 'n'), [1, 9]);
  const separator = fieldSwitch(instr, 'p')?.value;
  const linked = fieldSwitch(instr, 'h') !== undefined;
  const width = contentWidth(resolved, region.begin.paragraphId);

  const taken = new Set(opts.styleNames.keys());
  const styleOf = new Map<number, string>();
  const styles: StyleDefinition[] = [];
  const tocStyle = (level: number): string => {
    const known = styleOf.get(level);
    if (known !== undefined) return known;
    const want = `toc ${level}`;
    let id = [...opts.styleNames].find(([, name]) => normalizeName(name) === want)?.[0];
    if (id === undefined) {
      const def = tocStyleDefinition(level, (x) => taken.has(x), opts.normalStyleId);
      taken.add(def.id);
      styles.push(def);
      id = def.id;
    }
    styleOf.set(level, id);
    return id;
  };

  const paragraphs: ParagraphDraft[] = [];
  const bookmarks: TocPlan['bookmarks'] = [];
  for (const p of walkParagraphs(resolved)) {
    if (inToc.has(p.id)) continue;
    const level = select(p);
    if (level === undefined) continue;
    const text = headingText(p, keepTabs, keepBreaks);
    // Word 跳过没有字的标题（空的「标题 1」段落不会在目录里留一行空白）
    if (text === '') continue;

    let anchor = onParagraph.get(p.id)?.find((b) => /^_Toc/i.test(b));
    if (anchor === undefined) {
      do anchor = `_Toc${serial++}`;
      while (bookmarkNames.has(anchor));
      bookmarkNames.add(anchor);
      bookmarks.push({ paragraphId: p.id, name: anchor });
    }

    const pages = !(noPages !== undefined && level >= noPages[0] && level <= noPages[1]);
    const props: ParaProps = { styleId: tocStyle(level) };
    if (pages && separator === undefined && width !== undefined) {
      props.tabs = [{ pos: width, alignment: 'right', leader: 'dot' } satisfies TabStop];
    }
    const head: RunContent[] = [];
    const label = p.props.numbering.label;
    if (label !== undefined && label.text !== '') {
      head.push({ kind: 'text', text: label.text });
      if (label.suffix === 'tab') head.push({ kind: 'tab' });
      else if (label.suffix === 'space') head.push({ kind: 'text', text: ' ' });
    }
    head.push(...splitText(text));
    const runs: RunDraft[] = [];
    if (pages) {
      head.push(separator === undefined ? { kind: 'tab' } : { kind: 'text', text: separator });
      runs.push({ content: head }, ...pageRef(anchor, opts.pageText?.(p.id) ?? ''));
    } else runs.push({ content: head });
    if (linked) for (const r of runs) r.hyperlink = { anchor };
    paragraphs.push({ props, runs });
  }

  const entries = paragraphs.length;
  if (entries === 0)
    paragraphs.push({ props: {}, runs: [{ content: [{ kind: 'text', text: TOC_EMPTY_TEXT }] }] });
  return { ok: true, plan: { field: region.begin.runId, paragraphs, bookmarks, styles, entries } };
}

/** 嵌套的 `PAGEREF 书签 \h`：界桩各占一个 run，结果文字单独一个 run（域求值按 run 认结果区） */
function pageRef(anchor: string, page: string): RunDraft[] {
  return [
    { content: [{ kind: 'fieldChar', charType: 'begin' }] },
    { content: [{ kind: 'fieldInstruction', text: ` PAGEREF ${anchor} \\h ` }] },
    { content: [{ kind: 'fieldChar', charType: 'separate' }] },
    { content: [{ kind: 'text', text: page }] },
    { content: [{ kind: 'fieldChar', charType: 'end' }] },
  ];
}

/** `\w` / `\x` 保留下来的制表位与换行（标记成 U+0009 / U+000A）拆回片段 */
function splitText(text: string): RunContent[] {
  const out: RunContent[] = [];
  for (const piece of text.split(/([\t\n])/)) {
    if (piece === '') continue;
    if (piece === '\t') out.push({ kind: 'tab' });
    else if (piece === '\n') out.push({ kind: 'break', breakType: 'line' });
    else out.push({ kind: 'text', text: piece });
  }
  return out;
}

/** 标题段落进目录的文字：隐藏文字、域代码不进，域结果照进（与屏幕上看见的一致） */
function headingText(p: ResolvedParagraph, keepTabs: boolean, keepBreaks: boolean): string {
  let text = '';
  for (const run of p.runs) {
    if (run.props.hidden) continue;
    for (const c of run.content) {
      if (c.kind === 'text') text += c.text;
      else if (c.kind === 'tab') text += keepTabs ? '\t' : ' ';
      else if (c.kind === 'break' && c.breakType === 'line') text += keepBreaks ? '\n' : ' ';
      else if (c.kind === 'symbol') text += c.char;
      else if (c.kind === 'noBreakHyphen') text += '-';
    }
  }
  return text.trim();
}

/** `\o "1-3"` / `\n "2-2"` 这种级别范围；开关缺席答 undefined，没写值答 `all` */
function levelRange(
  sw: { value?: string } | undefined,
  all: readonly [number, number],
): readonly [number, number] | undefined {
  if (sw === undefined) return undefined;
  const m = /^\s*(\d)\s*(?:-\s*(\d))?\s*$/.exec(sw.value ?? '');
  if (m === null) return all;
  const lo = Number(m[1]);
  const hi = m[2] === undefined ? lo : Number(m[2]);
  return [Math.max(1, Math.min(lo, hi)), Math.min(9, Math.max(lo, hi))];
}

function normalizeName(name: string): string {
  return name.trim().toLowerCase().replace(/\s+/g, ' ');
}

/**
 * 段落 → 目录级别（1–9），不进目录答 undefined。`\t` 点名的样式优先；
 * 其余按大纲级别（`\o` 的范围，只写 `\u` 时 1–9）。三个开关都没写时 Word 按 `\o "1-9"` 处理
 */
function levelSelector(
  instr: FieldInstruction,
  styleNames: ReadonlyMap<string, string>,
): (p: ResolvedParagraph) => number | undefined {
  const byStyle = new Map<string, number>();
  const t = fieldSwitch(instr, 't')?.value;
  if (t !== undefined) {
    const parts = t.split(/[,;，；]/).map((s) => s.trim());
    for (let i = 0; i < parts.length; i += 2) {
      const name = parts[i];
      const level = Number(parts[i + 1] ?? '1');
      if (name === undefined || name === '' || !(level >= 1 && level <= 9)) continue;
      byStyle.set(normalizeName(name), level);
    }
  }
  const o = levelRange(fieldSwitch(instr, 'o'), [1, 9]);
  const u = fieldSwitch(instr, 'u') !== undefined;
  const outline = o ?? (u || t === undefined ? ([1, 9] as const) : undefined);
  return (p) => {
    if (byStyle.size) {
      const id = p.props.styleId;
      const hit = byStyle.get(normalizeName(styleNames.get(id) ?? id)) ?? byStyle.get(normalizeName(id));
      if (hit !== undefined) return hit;
    }
    if (outline === undefined) return undefined;
    const level = p.props.outlineLevel + 1;
    return level >= outline[0] && level <= outline[1] ? level : undefined;
  };
}

/** 所有 TOC 的结果区覆盖的段落（从 separate 那段到 end 那段）—— 目录自己的条目不能再收进目录 */
function tocResultParagraphs(resolved: ResolvedBody, fields: readonly FieldRegion[]): Set<NodeId> {
  const out = new Set<NodeId>();
  const spans = tocFields(fields).map((f) => [f.separate?.paragraphId, f.end?.paragraphId] as const);
  if (spans.length === 0) return out;
  let open = 0;
  for (const p of walkParagraphs(resolved)) {
    const starts = spans.filter(([s]) => s === p.id).length;
    open += starts;
    if (open > 0) out.add(p.id);
    open -= spans.filter(([, e]) => e === p.id).length;
  }
  return out;
}

/** 目录所在那一节的版心宽 —— 页码那个右对齐制表位就停在这儿 */
function contentWidth(resolved: ResolvedBody, paragraphId: NodeId): number | undefined {
  for (const section of resolved.sections) {
    for (const p of walkParagraphs({ sections: [section] })) {
      if (p.id !== paragraphId) continue;
      const { page, margin } = section.props;
      return page.width - margin.left - margin.right - margin.gutter;
    }
  }
  return undefined;
}

/** 新书签的编号接着文档里最大的 `_Toc` 编号往下数（Word 的是 9 位随机数，这里只要不撞） */
function nextTocSerial(names: ReadonlySet<string>): number {
  let max = 99999999;
  for (const n of names) {
    const m = /^_Toc(\d{1,15})$/i.exec(n);
    if (m !== null) max = Math.max(max, Number(m[1]));
  }
  return max + 1;
}
