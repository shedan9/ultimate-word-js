/**
 * `settings.xml` → 文档级设置。
 *
 * 这个部件很容易被当成「杂项」跳过，但里面有三样东西**直接决定坐标**：
 *
 * 1. `w:defaultTabStop` —— 没有显式制表位时每个 tab 走多远。不读就只能瞎猜
 * 2. `w:characterSpacingControl` —— **标点挤压**开不开。中文排版里行尾那个句号能不能压成
 *    半个字宽，整行的断行位置都会变
 * 3. `w:themeFontLang w:eastAsia` —— 主题字体里东亚脚本回退到哪个 script 标签。
 *    在此之前 `resolveThemeFont` 只能硬编码 `zh-CN`，日文文档会拿到简体中文字体
 *
 * 其余的 `w:compat` 开关原样收进一个字典：它们大多是「Word 某个版本的排版怪癖要不要复现」，
 * 现在还没到消费它们的阶段，但**丢了就再也不知道文档要求过什么**。
 */
import type { Twips } from '@uw/core';
import type { XmlDocument, XmlElement } from '@uw/ooxml';
import { attr, child, children } from '@uw/ooxml';
import { attrOf, enumVal, intVal, onOff } from './xml-values.ts';

export type CharacterSpacingControl =
  | 'doNotCompress'
  | 'compressPunctuation'
  | 'compressPunctuationAndJapaneseKana';

const SPACING_CONTROLS = [
  'doNotCompress',
  'compressPunctuation',
  'compressPunctuationAndJapaneseKana',
] as const;

export interface ThemeFontLang {
  /** 拉丁文语言（`w:val`） */
  latin: string;
  /** 东亚语言。主题字体的 EastAsia 回退按它选 script 标签 */
  eastAsia: string;
  /** 复杂文字语言 */
  bidi: string;
}

export interface DocumentSettings {
  /** 默认制表位间隔，twips。Word 中文模板是 420（= 21pt = 两个五号字宽） */
  defaultTabStop: Twips;
  characterSpacingControl: CharacterSpacingControl;
  themeFontLang: ThemeFontLang;
  /** `w:evenAndOddHeaders`：奇偶页用不同的页眉页脚 */
  evenAndOddHeaders: boolean;
  /** `w:mirrorMargins`：对称页边距（装订线在内侧） */
  mirrorMargins: boolean;
  /** `w:gutterAtTop`：装订线在上边而不是左边 */
  gutterAtTop: boolean;
  autoHyphenation: boolean;
  /** 断词区宽度，twips */
  hyphenationZone: Twips;
  /** 连续断词行数上限，0 表示不限 */
  consecutiveHyphenLimit: number;
  doNotHyphenateCaps: boolean;
  /**
   * `w:compatSetting name="compatibilityMode"`：文档要求按哪个版本的 Word 排版。
   * 15 = Word 2013 及以后。**这不是摆设** —— 12（2007）与 15 在行高、表格自动调整上有实打实的差异。
   * 0 表示文件里没写。
   */
  compatibilityMode: number;
  /**
   * `w:noLineBreaksAfter` —— 文档自定义的**尾禁则**字符集：这些字符之后不许断行，
   * 也就是「不能出现在行尾」的那批（`（「『【《` 之类）。
   *
   * 空串表示文件里没写，布局层用内建集合。写了就**整个替换**内建集合而不是追加 ——
   * 用户在 Word 里改禁则字符集时，界面上就是一个可编辑的完整列表。
   *
   * 规范允许按 `w:lang` 写多条（日文一套、中文一套）。这里把所有语言的合并成一串：
   * 判定只需要「这个字属不属于禁则集」，而一份文档不会同时按两套语言排版；
   * 真要分语言得先有能验证差异的真值样本，现在没有。
   */
  noLineBreaksAfter: string;
  /** `w:noLineBreaksBefore` —— **首禁则**：这些字符之前不许断行（`，。、）」` 等） */
  noLineBreaksBefore: string;
  /**
   * `w:compat` 底下所有开关，原样收着（元素名 → 是否开启）。
   *
   * 与排版真正相关、Phase 2 之后会来查的几个：
   * - `useFELayout` —— 用东亚排版规则（中文文档必开）
   * - `balanceSingleByteDoubleByteWidth` —— 半角字符按全角宽度对齐网格
   * - `doNotExpandShiftReturn` —— 两端对齐时，手动换行的那一行不拉伸
   * - `adjustLineHeightInTable` —— 表格内行高是否也按网格调整
   * - `spaceForUL` / `ulTrailSpace` —— 下划线是否延伸到行尾空格
   */
  compat: Record<string, boolean>;
}

export const DEFAULT_SETTINGS: DocumentSettings = {
  // Word 的出厂默认是 720（= 0.5 英寸）；中文模板会改成 420，但那是模板写进文件里的，
  // 不是默认值。这里必须给规范默认值，不能拿中文模板的习惯当默认
  defaultTabStop: 720,
  characterSpacingControl: 'doNotCompress',
  themeFontLang: { latin: '', eastAsia: '', bidi: '' },
  evenAndOddHeaders: false,
  mirrorMargins: false,
  gutterAtTop: false,
  autoHyphenation: false,
  hyphenationZone: 360,
  consecutiveHyphenLimit: 0,
  doNotHyphenateCaps: false,
  noLineBreaksAfter: '',
  noLineBreaksBefore: '',
  compatibilityMode: 0,
  compat: {},
};

export function parseSettings(doc: XmlDocument | undefined): DocumentSettings {
  if (doc === undefined) return structuredClone(DEFAULT_SETTINGS);
  const root = doc.root;
  const d = DEFAULT_SETTINGS;
  const compatEl = child(root, 'w:compat');

  return {
    defaultTabStop: intVal(root, 'w:defaultTabStop') ?? d.defaultTabStop,
    characterSpacingControl:
      enumVal(valOfRoot(root, 'w:characterSpacingControl'), SPACING_CONTROLS) ?? d.characterSpacingControl,
    themeFontLang: parseThemeFontLang(child(root, 'w:themeFontLang')),
    evenAndOddHeaders: onOff(root, 'w:evenAndOddHeaders') ?? d.evenAndOddHeaders,
    mirrorMargins: onOff(root, 'w:mirrorMargins') ?? d.mirrorMargins,
    gutterAtTop: onOff(root, 'w:gutterAtTop') ?? d.gutterAtTop,
    autoHyphenation: onOff(root, 'w:autoHyphenation') ?? d.autoHyphenation,
    hyphenationZone: intVal(root, 'w:hyphenationZone') ?? d.hyphenationZone,
    consecutiveHyphenLimit: intVal(root, 'w:consecutiveHyphenLimit') ?? d.consecutiveHyphenLimit,
    doNotHyphenateCaps: onOff(root, 'w:doNotHyphenateCaps') ?? d.doNotHyphenateCaps,
    noLineBreaksAfter: mergeKinsoku(root, 'w:noLineBreaksAfter'),
    noLineBreaksBefore: mergeKinsoku(root, 'w:noLineBreaksBefore'),
    compatibilityMode: compatSettingInt(compatEl, 'compatibilityMode') ?? d.compatibilityMode,
    compat: parseCompat(compatEl),
  };
}

function valOfRoot(root: XmlElement, name: string): string | undefined {
  return attrOf(child(root, name), 'w:val');
}

/**
 * 把各语言的禁则字符集合并成一串并去重。
 *
 * 按码点去重而不是按 UTF-16 码元：禁则集里理论上可以出现补充平面的字符，
 * 按码元切会把代理对拆成两个半个字符，之后所有比较都不会再命中。
 */
function mergeKinsoku(root: XmlElement, name: string): string {
  const seen = new Set<string>();
  for (const el of children(root, name)) {
    for (const ch of attr(el, 'w:val') ?? '') seen.add(ch);
  }
  return [...seen].join('');
}

function parseThemeFontLang(el: XmlElement | undefined): ThemeFontLang {
  return {
    latin: attrOf(el, 'w:val') ?? '',
    eastAsia: attrOf(el, 'w:eastAsia') ?? '',
    bidi: attrOf(el, 'w:bidi') ?? '',
  };
}

/**
 * `w:compat` 的子元素分两种：光秃秃的开关（`<w:useFELayout/>`）和
 * `<w:compatSetting w:name="..." w:val="..."/>`。前者按 ST_OnOff 收进字典，
 * 后者按名字收 —— 值不是布尔的（如 compatibilityMode=15）另行取。
 */
const ON_OFF_VALUES = new Set(['0', '1', 'true', 'false', 'on', 'off']);

function parseCompat(compatEl: XmlElement | undefined): Record<string, boolean> {
  const out: Record<string, boolean> = {};
  if (compatEl === undefined) return out;
  for (const el of children(compatEl)) {
    if (el.name === 'w:compatSetting') {
      const name = attr(el, 'w:name');
      const val = attr(el, 'w:val');
      // 只收真正是开关的：compatibilityMode 的值是 15 这种版本号，塞进布尔字典里
      // 会读成「true」，误导得很。它有自己的字段
      if (name !== undefined && val !== undefined && ON_OFF_VALUES.has(val)) {
        out[name] = val !== '0' && val !== 'false' && val !== 'off';
      }
      continue;
    }
    // 去掉 w: 前缀，调用方按裸名字查（`compat.useFELayout`），不必记着前缀
    const key = el.name.startsWith('w:') ? el.name.slice(2) : el.name;
    const val = attr(el, 'w:val');
    out[key] = val === undefined ? true : val !== '0' && val !== 'false' && val !== 'off';
  }
  return out;
}

function compatSettingInt(compatEl: XmlElement | undefined, name: string): number | undefined {
  if (compatEl === undefined) return undefined;
  for (const el of children(compatEl, 'w:compatSetting')) {
    if (attr(el, 'w:name') !== name) continue;
    const n = Number.parseInt(attr(el, 'w:val') ?? '', 10);
    return Number.isNaN(n) ? undefined : n;
  }
  return undefined;
}
