import { describe, expect, it } from 'vitest';
import type { DateTimeParts } from './date-format.ts';
import { formatDateTime, localDateTimeParts } from './date-format.ts';

/** 2026-09-29 星期二 08:05:03 */
const T: DateTimeParts = { year: 2026, month: 9, day: 29, hour: 8, minute: 5, second: 3, weekday: 2 };

describe('formatDateTime', () => {
  it('数字记号：宽度看连写几次，M 是月、m 是分（区分大小写）', () => {
    expect(formatDateTime(T, 'yyyy-MM-dd HH:mm:ss')).toBe('2026-09-29 08:05:03');
    expect(formatDateTime(T, 'yy/M/d H:m')).toBe('26/9/29 8:5');
  });

  it('中文格式串：汉字原样照抄，单引号里的也照抄', () => {
    expect(formatDateTime(T, "yyyy'年'M'月'd'日'")).toBe('2026年9月29日');
    expect(formatDateTime(T, 'yyyy年M月d日')).toBe('2026年9月29日');
    expect(formatDateTime(T, "'at' h")).toBe('at 8');
    expect(formatDateTime(T, "h''mm")).toBe("8'05");
  });

  it('12 小时制与 AM/PM；0 点与 12 点显示 12', () => {
    expect(formatDateTime(T, 'h:mm AM/PM')).toBe('8:05 AM');
    expect(formatDateTime({ ...T, hour: 0 }, 'h am/pm')).toBe('12 am');
    expect(formatDateTime({ ...T, hour: 12 }, 'hh AM/PM')).toBe('12 PM');
  });

  it('月名与星期：英文格式串出英文，带汉字的出中文', () => {
    expect(formatDateTime(T, 'dddd, MMMM d, yyyy')).toBe('Tuesday, September 29, 2026');
    expect(formatDateTime(T, 'ddd MMM')).toBe('Tue Sep');
    expect(formatDateTime(T, 'yyyy年M月d日dddd')).toBe('2026年9月29日星期二');
  });

  it('中文版的 EEEE年O月A日：年份逐位念、月日按中文计数', () => {
    expect(formatDateTime(T, 'EEEE年O月A日')).toBe('二〇二六年九月二十九日');
    expect(formatDateTime({ ...T, month: 10, day: 10 }, 'O月A日')).toBe('十月十日');
    expect(formatDateTime({ ...T, month: 12, day: 1 }, 'O月A日')).toBe('十二月一日');
  });

  it('localDateTimeParts 取本机时区的分量', () => {
    const d = new Date(2026, 8, 29, 23, 59, 1);
    expect(localDateTimeParts(d)).toEqual({
      year: 2026,
      month: 9,
      day: 29,
      hour: 23,
      minute: 59,
      second: 1,
      weekday: 2,
    });
  });
});
