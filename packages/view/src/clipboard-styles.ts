/**
 * 剪贴板中的样式只读原文，不挂到宿主页面；CSSOM 会丢掉 mso-* 声明。
 * Word 常用的标签、类、id 及其简单组合按优先级级联；复杂选择器和整块 at-rule 跳过。
 * 不读取外部样式，不注册字体，也不把 URL 或其他 CSS 内容交给渲染器。
 */
interface Declaration {
  value: string;
  important: boolean;
}
type Specificity = readonly [number, number, number];
interface Rule {
  selector: string;
  specificity: Specificity;
  declarations: Map<string, Declaration>;
}

/** 只归一粘贴层会消费的长短属性，避免类中的简写永远盖住后来的内联长属性。 */
const PROPERTY_NAMES: ReadonlyMap<string, string> = new Map([
  ['background', 'background-color'],
  ['text-decoration', 'text-decoration-line'],
  ['text-emphasis', 'text-emphasis-style'],
  ['page-break-before', 'break-before'],
]);

/** 注释里的分隔符不能截断规则；引号里的注释标记则是字体名等值的一部分。 */
function withoutComments(text: string): string {
  let out = '';
  let quote = '';
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '\\') {
      out += text.slice(i, i + 2);
      i++;
    } else if (quote) {
      out += ch;
      if (ch === quote) quote = '';
    } else if (ch === '"' || ch === "'") {
      quote = ch;
      out += ch;
    } else if (ch === '/' && text[i + 1] === '*') {
      const end = text.indexOf('*/', i + 2);
      if (end < 0) break;
      out += ' ';
      i = end + 1;
    } else out += ch;
  }
  return out;
}

/** 括号与引号里的分号不是声明边界；同一属性先比较 important，再取最后一项。 */
function declarations(text: string): Map<string, Declaration> {
  const clean = withoutComments(text);
  const out = new Map<string, Declaration>();
  let quote = '';
  let parens = 0;
  let start = 0;
  const flush = (end: number) => {
    const decl = clean.slice(start, end);
    const colon = decl.indexOf(':');
    if (colon < 1) return;
    const property = decl.slice(0, colon).trim().toLowerCase();
    const name = PROPERTY_NAMES.get(property) ?? property;
    const raw = decl.slice(colon + 1).trim();
    const important = /!\s*important\s*$/i.test(raw);
    const value = raw.replace(/\s*!\s*important\s*$/i, '').trim();
    if (value && (important || !out.get(name)?.important)) out.set(name, { value, important });
  };
  for (let i = 0; i < clean.length; i++) {
    const ch = clean[i];
    if (ch === '\\') i++;
    else if (quote) {
      if (ch === quote) quote = '';
    } else if (ch === '"' || ch === "'") quote = ch;
    else if (ch === '(') parens++;
    else if (ch === ')') parens = Math.max(0, parens - 1);
    else if (ch === ';' && parens === 0) {
      flush(i);
      start = i + 1;
    }
  }
  flush(clean.length);
  return out;
}

/** 不借浏览器计算样式，避免未知的选择器、媒体条件或外部资源影响粘贴结果。 */
function specificity(selector: string): Specificity | undefined {
  if (!/^(?:[a-z][\w-]*|\*)?(?:[.#][\w\u0080-\uffff-]+)*$/i.test(selector) || !selector) return undefined;
  return [
    (selector.match(/#/g) ?? []).length,
    (selector.match(/\./g) ?? []).length,
    /^[a-z]/i.test(selector) ? 1 : 0,
  ];
}

function rules(text: string): Rule[] {
  const clean = withoutComments(text)
    .replace(/^\s*<!--/, '')
    .replace(/-->\s*$/, '');
  const out: Rule[] = [];
  let quote = '';
  let parens = 0;
  let depth = 0;
  let start = 0;
  let body = 0;
  let header = '';
  for (let i = 0; i < clean.length; i++) {
    const ch = clean[i];
    if (ch === '\\') i++;
    else if (quote) {
      if (ch === quote) quote = '';
    } else if (ch === '"' || ch === "'") quote = ch;
    else if (ch === '(') parens++;
    else if (ch === ')') parens = Math.max(0, parens - 1);
    else if (parens === 0 && ch === '{') {
      if (depth === 0) {
        header = clean.slice(start, i).trim();
        body = i + 1;
      }
      depth++;
    } else if (parens === 0 && ch === '}') {
      if (depth > 0 && --depth === 0 && !header.startsWith('@')) {
        const props = declarations(clean.slice(body, i));
        for (const selector of header.split(',').map((s) => s.trim())) {
          const weight = specificity(selector);
          if (weight) out.push({ selector, specificity: weight, declarations: props });
        }
      }
      if (depth === 0) start = i + 1;
    } else if (depth === 0 && parens === 0 && ch === ';') start = i + 1;
  }
  return out;
}

function compare(a: Specificity, b: Specificity): number {
  return a[0] - b[0] || a[1] - b[1] || a[2] - b[2];
}

/** 每份剪贴板解析一次样式表；规则的顺序包含 style 元素之间的文档顺序。 */
export function clipboardStyleReader(doc: Document, trusted: boolean): (el: Element) => Map<string, string> {
  const sheet = trusted
    ? [...doc.querySelectorAll('style')].flatMap((el) => {
        const type = el.getAttribute('type')?.trim().toLowerCase();
        const media = el.getAttribute('media')?.trim().toLowerCase();
        if (type && type !== 'text/css') return [];
        if (media && !media.split(',').some((part) => /^(all|screen)$/.test(part.trim()))) return [];
        return rules(el.textContent ?? '');
      })
    : [];
  return (el) => {
    const chosen = new Map<string, Declaration & { specificity: Specificity; inline: boolean }>();
    const apply = (props: Map<string, Declaration>, weight: Specificity, inline: boolean) => {
      for (const [name, decl] of props) {
        const previous = chosen.get(name);
        if (
          !previous ||
          (decl.important !== previous.important
            ? decl.important
            : inline !== previous.inline
              ? inline
              : compare(weight, previous.specificity) >= 0)
        )
          chosen.set(name, { ...decl, specificity: weight, inline });
      }
    };
    for (const rule of sheet) {
      try {
        if (el.matches(rule.selector)) apply(rule.declarations, rule.specificity, false);
      } catch {
        // 未实现的转义或非法标识符只忽略本条，不让整个粘贴失败。
      }
    }
    apply(declarations(el.getAttribute('style') ?? ''), [0, 0, 0], true);
    return new Map([...chosen].map(([name, decl]) => [name, decl.value]));
  };
}
