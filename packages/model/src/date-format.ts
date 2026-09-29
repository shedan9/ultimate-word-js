/**
 * 日期时间 + `\@` 格式串（date-time picture）→ 文字。DATE / TIME 域用它。
 *
 * 纯函数：时间以**已经换算到本地**的各个分量传进来（`DateTimeParts`），不碰 `Date`
 * 与时区 —— 换算是调用方（门面）的事，这样单测不必关心跑在哪个时区，
 * 离线工具也能给一个固定的时刻复现同一份布局。
 *
 * 格式串的记号照 ECMA-376 §17.16.4.1，**区分大小写**：`M` 是月、`m` 是分，`H` 是 24 小时、
 * `h` 是 12 小时；单引号里的原样照抄，其余不认识的字符（`年` `/` `-` `:`）也原样照抄。
 *
 * ── 未标定（没有 Word 样本）──────────────────────────────────────────────────
 * 1. 中文版 Word「日期和时间」对话框里「二〇二六年九月二十九日」那一项写出来的是
 *    `\@ "EEEE年O月A日"`：`E` 系列 = 年份逐位念（〇一二…）、`O` = 月、`A` = 日（中文计数）。
 *    这三个字母不在规范里，读法按对话框里的预览写的
 * 2. `MMMM` / `MMM` / `dddd` / `ddd` 的**语言**：Word 按域所在 run 的语言出（中文版「九月」「星期二」），
 *    这里按格式串里有没有汉字猜 —— 中文格式串里几乎总带「年」「月」
 * 钉死办法：一份 docx，放 `DATE \@` 的十几种写法（含 `EEEE年O月A日`、`dddd`、`MMMM`、中英文 run），
 * 更新后读 `document.xml` 里的结果文字 —— 结果就存在文件里，不必导 PDF。
 */

/** 本地时间的各个分量（纯数据，可结构化克隆） */
export interface DateTimeParts {
  year: number;
  /** 1–12 */
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
  /** 0 = 星期日 … 6 = 星期六，与 `Date.prototype.getDay()` 一致 */
  weekday: number;
}

/** `Date` → 它在**本机时区**的各个分量。Word 显示的是本机时钟，不是 UTC */
export function localDateTimeParts(date: Date): DateTimeParts {
  return {
    year: date.getFullYear(),
    month: date.getMonth() + 1,
    day: date.getDate(),
    hour: date.getHours(),
    minute: date.getMinutes(),
    second: date.getSeconds(),
    weekday: date.getDay(),
  };
}

const EN_MONTHS = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
] as const;
const EN_WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'] as const;
const CN_WEEKDAYS = ['日', '一', '二', '三', '四', '五', '六'] as const;
const CN_DIGIT = ['〇', '一', '二', '三', '四', '五', '六', '七', '八', '九'] as const;

/** 1–99 的中文计数读法（九、十、十二、二十九）。月与日用不到更大的数 */
function chineseCount(n: number): string {
  if (n < 10) return CN_DIGIT[n] ?? String(n);
  const tens = Math.floor(n / 10);
  const ones = n % 10;
  return `${tens === 1 ? '' : (CN_DIGIT[tens] ?? '')}十${ones === 0 ? '' : (CN_DIGIT[ones] ?? '')}`;
}

const pad2 = (n: number) => String(n).padStart(2, '0');

/**
 * 把一个日期时间按 `\@` 格式串排成文字。
 *
 * 认不出的记号不会抛：它们原样留在结果里（与 Word 一样 —— 格式串里的汉字、标点本来就是这么进结果的）。
 */
export function formatDateTime(t: DateTimeParts, picture: string): string {
  const chinese = /[一-鿿]/u.test(picture);
  let out = '';
  let i = 0;
  while (i < picture.length) {
    const ch = picture.charAt(i);
    if (ch === "'") {
      // 单引号里的原样照抄；`''` 是一个单引号本身
      const close = picture.indexOf("'", i + 1);
      if (close === i + 1) {
        out += "'";
        i += 2;
        continue;
      }
      out += close < 0 ? picture.slice(i + 1) : picture.slice(i + 1, close);
      i = close < 0 ? picture.length : close + 1;
      continue;
    }
    if (picture.startsWith('AM/PM', i) || picture.startsWith('am/pm', i)) {
      const pm = t.hour >= 12;
      out += ch === 'A' ? (pm ? 'PM' : 'AM') : pm ? 'pm' : 'am';
      i += 5;
      continue;
    }
    // 同一个字母连写几次决定宽度：M / MM / MMM / MMMM
    let n = 1;
    while (picture.charAt(i + n) === ch) n++;
    const token = tokenText(t, ch, n, chinese);
    if (token === undefined) {
      out += ch;
      i++;
      continue;
    }
    out += token;
    i += n;
  }
  return out;
}

function tokenText(t: DateTimeParts, ch: string, n: number, chinese: boolean): string | undefined {
  switch (ch) {
    case 'y':
    case 'Y':
      // `yy` 两位，其余一律四位 —— Word 对 `y` / `yyy` 的处理没有说法，给全年最不会误导
      return n === 2 ? pad2(t.year % 100) : String(t.year);
    case 'M':
      if (n >= 4) return chinese ? `${chineseCount(t.month)}月` : (EN_MONTHS[t.month - 1] ?? '');
      if (n === 3) return chinese ? `${chineseCount(t.month)}月` : (EN_MONTHS[t.month - 1] ?? '').slice(0, 3);
      return n === 2 ? pad2(t.month) : String(t.month);
    case 'd':
      if (n >= 4) return chinese ? `星期${CN_WEEKDAYS[t.weekday] ?? ''}` : (EN_WEEKDAYS[t.weekday] ?? '');
      if (n === 3)
        return chinese ? `周${CN_WEEKDAYS[t.weekday] ?? ''}` : (EN_WEEKDAYS[t.weekday] ?? '').slice(0, 3);
      return n === 2 ? pad2(t.day) : String(t.day);
    case 'H':
      return n >= 2 ? pad2(t.hour) : String(t.hour);
    case 'h': {
      const h12 = t.hour % 12 === 0 ? 12 : t.hour % 12;
      return n >= 2 ? pad2(h12) : String(h12);
    }
    case 'm':
      return n >= 2 ? pad2(t.minute) : String(t.minute);
    case 's':
    case 'S':
      return n >= 2 ? pad2(t.second) : String(t.second);
    // 中文版 Word 的三个字母，见文件头「未标定」第 1 条
    case 'E':
      return [...String(t.year)].map((d) => CN_DIGIT[Number(d)] ?? d).join('');
    case 'O':
      return chineseCount(t.month);
    case 'A':
      return chineseCount(t.day);
    default:
      return undefined;
  }
}
