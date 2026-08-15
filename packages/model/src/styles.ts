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
import type { ParaProps, RunProps } from './props.ts';
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
}

export interface StyleSheet {
  /** `docDefaults` —— 级联的第一层 */
  defaults: { paraProps: ParaProps; runProps: RunProps };
  byId(id: string): Style | undefined;
  /** 默认段落样式 id（通常是 Normal 那个）。找不到返回空串 */
  defaultParagraphStyleId(): string;
  defaultCharacterStyleId(): string;
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
  };
}
