/**
 * 样式表（`styles.xml`）。
 *
 * 两件事：把每个 `w:style` 翻译成属性对象，以及把 `basedOn` 链**展平成数组**。
 * 链的方向要想清楚：`basedOn` 指向父样式，但**应用顺序是从祖先到自己** ——
 * 祖先先铺，后代覆盖。所以展平后要反过来。搞反的话「标题 1 基于正文」会变成
 * 正文覆盖标题，字号全错。
 *
 * `basedOn` 成环是**内容问题不是结构错误**（架构 §10）：记一条诊断，
 * 把环截断继续跑。用户要的是看到文档。
 */
import type { DiagnosticSink } from '@uw/core';
import type { XmlDocument, XmlElement } from '@uw/ooxml';
import { attr, child, children } from '@uw/ooxml';
import { parseParaProps, parseRunProps } from './parse-props.ts';
import { parseCellProps, parseRowProps, parseTableProps } from './parse-table-props.ts';
import type { ParaProps, RunProps } from './props.ts';
import type {
  CellProps,
  RowProps,
  TableProps,
  TableStyleOverride,
  TableStyleOverrideType,
} from './table-props.ts';
import { CONDITIONAL_ORDER } from './table-props.ts';
import { valOf } from './xml-values.ts';

export type StyleType = 'paragraph' | 'character' | 'table' | 'numbering';

export interface Style {
  id: string;
  type: StyleType;
  /** `w:name`，界面上显示的名字（`Normal` / `heading 1`），与 id 是两回事 */
  name: string;
  basedOn: string | undefined;
  /** `w:next`：按回车后下一段用哪个样式。编辑态才用得到 */
  next: string | undefined;
  /** 这个类型的默认样式（`w:default="1"`） */
  isDefault: boolean;
  paraProps: ParaProps;
  runProps: RunProps;
  /** `w:tblPr` / `w:trPr` / `w:tcPr`。只有 `type="table"` 的样式才可能非空 */
  tableProps: TableProps;
  rowProps: RowProps;
  cellProps: CellProps;
  /**
   * `w:tblStylePr`：条件格式（首行 / 末列 / 隔行带…）。
   * 一个类型最多一份，重复出现按**后者胜**收进 Map —— 与其它「后面覆盖前面」一致。
   */
  conditional: Map<TableStyleOverrideType, TableStyleOverride>;
}

export interface StyleSheet {
  /** `docDefaults` —— 级联的第一层 */
  defaults: { paraProps: ParaProps; runProps: RunProps };
  byId(id: string): Style | undefined;
  /** 默认段落样式 id（通常是 Normal 那个）。找不到返回空串 */
  defaultParagraphStyleId(): string;
  defaultCharacterStyleId(): string;
  /**
   * 默认表格样式（通常是 `Normal Table`）。
   *
   * 它**不是**摆设：Word 模板里单元格默认的左右边距 108 twips 就写在这份样式的
   * `w:tblCellMar` 里。没写 `w:tblStyle` 的表格照样吃它 —— 与段落吃 `Normal` 同构。
   * 把 108 硬编码成兜底常数是错的，那样用户改了模板我们也跟不上。
   */
  defaultTableStyleId(): string;
  /**
   * 展平后的样式链，**从祖先到自己**，可以直接按序 apply。
   * 结果带缓存：一份公文里同一个样式会被几百个段落问到。
   */
  chainOf(styleId: string | undefined): Style[];
  all(): Style[];
}

const STYLE_TYPES: readonly StyleType[] = ['paragraph', 'character', 'table', 'numbering'];

export function parseStyles(doc: XmlDocument | undefined, diagnostics: DiagnosticSink): StyleSheet {
  const styles = new Map<string, Style>();
  let defaults = { paraProps: {} as ParaProps, runProps: {} as RunProps };

  if (doc !== undefined) {
    const docDefaults = child(doc.root, 'w:docDefaults');
    if (docDefaults !== undefined) {
      // 多一层包装：pPrDefault > pPr，rPrDefault > rPr
      const pPrDefault = child(docDefaults, 'w:pPrDefault');
      const rPrDefault = child(docDefaults, 'w:rPrDefault');
      defaults = {
        paraProps: parseParaProps(pPrDefault && child(pPrDefault, 'w:pPr')),
        runProps: parseRunProps(rPrDefault && child(rPrDefault, 'w:rPr')),
      };
    }
    for (const el of children(doc.root, 'w:style')) {
      const style = parseStyle(el);
      if (style !== undefined) styles.set(style.id, style);
    }
  }

  const defaultIdOf = (type: StyleType): string => {
    for (const s of styles.values()) {
      if (s.type === type && s.isDefault) return s.id;
    }
    return '';
  };
  const defaultPara = defaultIdOf('paragraph');
  const defaultChar = defaultIdOf('character');
  const defaultTable = defaultIdOf('table');

  const chainCache = new Map<string, Style[]>();

  const chainOf = (styleId: string | undefined): Style[] => {
    if (styleId === undefined || styleId === '') return [];
    const cached = chainCache.get(styleId);
    if (cached !== undefined) return cached;

    // 从自己往上爬到根，seen 同时充当成环检测
    const upward: Style[] = [];
    const seen = new Set<string>();
    let cursor: string | undefined = styleId;
    while (cursor !== undefined && cursor !== '') {
      if (seen.has(cursor)) {
        diagnostics.warn(
          'style-cycle',
          `样式 basedOn 成环，已在 ${cursor} 处截断：${[...seen].join(' → ')} → ${cursor}`,
          { part: '/word/styles.xml', path: `w:style[@w:styleId='${styleId}']` },
        );
        break;
      }
      seen.add(cursor);
      const style: Style | undefined = styles.get(cursor);
      if (style === undefined) {
        // 引用了不存在的样式：同样是内容问题，跳过继续
        diagnostics.warn('style-missing', `样式 ${cursor} 不存在，已跳过`, {
          part: '/word/styles.xml',
          path: `w:style[@w:styleId='${styleId}']`,
        });
        break;
      }
      upward.push(style);
      cursor = style.basedOn;
    }

    // 爬的时候是「自己 → 祖先」，应用要「祖先 → 自己」
    const chain = upward.reverse();
    chainCache.set(styleId, chain);
    return chain;
  };

  return {
    defaults,
    byId: (id) => styles.get(id),
    defaultParagraphStyleId: () => defaultPara,
    defaultCharacterStyleId: () => defaultChar,
    defaultTableStyleId: () => defaultTable,
    chainOf,
    all: () => [...styles.values()],
  };
}

function parseStyle(el: XmlElement): Style | undefined {
  const id = attr(el, 'w:styleId');
  if (id === undefined) return undefined;
  const type = attr(el, 'w:type');
  return {
    id,
    type: STYLE_TYPES.includes(type as StyleType) ? (type as StyleType) : 'paragraph',
    name: valOf(el, 'w:name') ?? '',
    basedOn: valOf(el, 'w:basedOn'),
    next: valOf(el, 'w:next'),
    // w:default 挂在属性上而不是子元素里，用 onOff 读不到
    isDefault: attr(el, 'w:default') === '1' || attr(el, 'w:default') === 'true',
    paraProps: parseParaProps(child(el, 'w:pPr')),
    runProps: parseRunProps(child(el, 'w:rPr')),
    tableProps: parseTableProps(child(el, 'w:tblPr')),
    rowProps: parseRowProps(child(el, 'w:trPr')),
    cellProps: parseCellProps(child(el, 'w:tcPr')),
    conditional: parseConditional(el),
  };
}

const OVERRIDE_TYPES: readonly TableStyleOverrideType[] = CONDITIONAL_ORDER;

/**
 * `w:tblStylePr` 一份份收进 Map。
 *
 * 认不出的 `w:type` **丢掉而不是报诊断**：这里没有 sink，而且它的后果只是
 * 少一层格式，不会丢内容。真正需要出声的未知元素在 parse-body.ts 那条路上。
 */
function parseConditional(el: XmlElement): Map<TableStyleOverrideType, TableStyleOverride> {
  const out = new Map<TableStyleOverrideType, TableStyleOverride>();
  for (const pr of children(el, 'w:tblStylePr')) {
    const type = attr(pr, 'w:type');
    if (type === undefined || !OVERRIDE_TYPES.includes(type as TableStyleOverrideType)) continue;
    out.set(type as TableStyleOverrideType, {
      paraProps: parseParaProps(child(pr, 'w:pPr')),
      runProps: parseRunProps(child(pr, 'w:rPr')),
      tableProps: parseTableProps(child(pr, 'w:tblPr')),
      rowProps: parseRowProps(child(pr, 'w:trPr')),
      cellProps: parseCellProps(child(pr, 'w:tcPr')),
    });
  }
  return out;
}
