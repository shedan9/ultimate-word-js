import { describe, expect, it } from 'vitest';
import { pagesToRender } from './virtual-pages.ts';

describe('可见页与预绘制范围', () => {
  it('保留视口前后两页，首尾不越界', () => {
    expect([...pagesToRender(10, [0, 1], 2)]).toEqual([0, 1, 2, 3]);
    expect([...pagesToRender(10, [9], 2)]).toEqual([7, 8, 9]);
  });
  it('滚动到中间后释放远处页面，多个可见区间不填满空隙', () => {
    expect([...pagesToRender(100, [50], 1)]).toEqual([49, 50, 51]);
    expect([...pagesToRender(100, [1, 80], 0)]).toEqual([1, 80]);
  });
  it('整个视图离屏时不保留绘制内容，忽略过期的 observer 下标', () => {
    expect([...pagesToRender(10, [], 2)]).toEqual([]);
    expect([...pagesToRender(0, [0], 2)]).toEqual([]);
    expect([...pagesToRender(2, [-1, 2], 2)]).toEqual([]);
  });
});
