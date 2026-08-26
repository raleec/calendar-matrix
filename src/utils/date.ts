/** Number of days in the given zero-indexed `month` (0 = January) of `year`. */
export function daysInMonth(year: number, month: number): number {
  return new Date(year, month + 1, 0).getDate()
}
