/**
 * 主题字体（`theme1.xml` 的 `a:fontScheme`）。
 *
 * 为什么这是必做项而不是锦上添花：Word 的 `docDefaults` 里正文字体写的是
 * `w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"` —— **一个真实字体名都没有**。
 * 不解析主题，整份文档的默认字体就是空的，度量无从谈起。
 */
import type { XmlDocument, XmlElement } from '@uw/ooxml';
import { attr, child, children } from '@uw/ooxml';

export interface FontScheme {
  /** `a:latin`，拉丁字体 */
  latin: string;
  /** `a:ea`，东亚字体。**经常是空串**，见 resolveThemeFont */
  eastAsia: string;
  /** `a:cs`，复杂文字字体 */
  cs: string;
  /** `a:font script="Hans"` 之类的按脚本回退表 */
  byScript: Record<string, string>;
}

export interface Theme {
  major: FontScheme;
  minor: FontScheme;
}

const EMPTY_SCHEME: FontScheme = { latin: '', eastAsia: '', cs: '', byScript: {} };

/** 没有 theme1.xml 是合法的（不是所有 docx 都有），此时所有主题引用解析成空串 */
export const EMPTY_THEME: Theme = { major: EMPTY_SCHEME, minor: EMPTY_SCHEME };

export function parseTheme(doc: XmlDocument): Theme {
  const elements = child(doc.root, 'a:themeElements');
  const scheme = elements && child(elements, 'a:fontScheme');
  if (!scheme) return EMPTY_THEME;
  return {
    major: parseScheme(child(scheme, 'a:majorFont')),
    minor: parseScheme(child(scheme, 'a:minorFont')),
  };
}

function parseScheme(el: XmlElement | undefined): FontScheme {
  if (el === undefined) return EMPTY_SCHEME;
  const byScript: Record<string, string> = {};
  for (const f of children(el, 'a:font')) {
    const script = attr(f, 'script');
    const typeface = attr(f, 'typeface');
    if (script !== undefined && typeface !== undefined && typeface !== '') byScript[script] = typeface;
  }
  return {
    latin: typefaceOf(el, 'a:latin'),
    eastAsia: typefaceOf(el, 'a:ea'),
    cs: typefaceOf(el, 'a:cs'),
    byScript,
  };
}

function typefaceOf(parent: XmlElement, name: string): string {
  const el = child(parent, name);
  return el === undefined ? '' : (attr(el, 'typeface') ?? '');
}

/**
 * `w:lang w:eastAsia` → 主题里的脚本标签。
 * 只列真的会碰到的四个，其余一律按简体中文处理 —— 这个库的定位就是中文公文。
 */
function scriptTagOf(langEastAsia: string): string {
  const lang = langEastAsia.toLowerCase();
  if (lang.startsWith('ja')) return 'Jpan';
  if (lang.startsWith('ko')) return 'Hang';
  if (lang === 'zh-tw' || lang === 'zh-hk' || lang === 'zh-mo') return 'Hant';
  return 'Hans';
}

/** 主题字体引用的取值：`minorHAnsi` / `majorEastAsia` / `minorBidi` … */
export type ThemeFontRef = string;

/**
 * 主题引用 → 真实字体名。
 *
 * 那个坑：**`a:ea` 在中文版 Office 的主题里是空串**（实测 gongwen-01.docx 的
 * `<a:ea typeface=""/>`），东亚字体真正写在 `<a:font script="Hans" typeface="等线"/>` 里。
 * 直接取 `a:ea` 会得到空字符串，整份公文的默认中文字体就没了 —— 而且是静默的，
 * 直到度量阶段才炸。所以 ea 为空时必须按语言回退到 byScript。
 */
export function resolveThemeFont(theme: Theme, ref: ThemeFontRef, langEastAsia = 'zh-CN'): string {
  const isMajor = ref.startsWith('major');
  const scheme = isMajor ? theme.major : theme.minor;
  // 'major' 与 'minor' 都是 5 个字符，剥掉前缀剩下 Ascii / HAnsi / EastAsia / Bidi
  const kind = isMajor || ref.startsWith('minor') ? ref.slice(5) : ref;
  switch (kind) {
    case 'EastAsia':
      return scheme.eastAsia !== '' ? scheme.eastAsia : (scheme.byScript[scriptTagOf(langEastAsia)] ?? '');
    case 'Bidi':
      return scheme.cs;
    // Ascii / HAnsi 都取 a:latin
    default:
      return scheme.latin;
  }
}
