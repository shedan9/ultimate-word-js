/** 按数组位置扩展可见页，不使用可能重新起算的显示页码。 */
export function pagesToRender(count: number, visible: Iterable<number>, overscan: number): Set<number> {
  const result = new Set<number>();
  for (const index of visible) {
    if (!Number.isInteger(index) || index < 0 || index >= count) continue;
    for (let i = Math.max(0, index - overscan); i <= Math.min(count - 1, index + overscan); i++) {
      result.add(i);
    }
  }
  return result;
}
