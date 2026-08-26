/** Number of days in the given zero-indexed `month` (0 = January) of `year`. */
export function daysInMonth(year: number, month: number): number {
  return new Date(year, month + 1, 0).getDate()
}

const DAY_LABEL_FORMATTER = new Intl.DateTimeFormat('en-US', {
  day: 'numeric',
  month: 'long',
})

/** Formats a date as e.g. "14 July" for use in cell `aria-label`s. */
export function formatDayLabel(
  year: number,
  month: number,
  day: number,
): string {
  const [monthPart, dayPart] = DAY_LABEL_FORMATTER.formatToParts(
    new Date(year, month, day),
  ).reduce<[string, string]>(
    (parts, part) => {
      if (part.type === 'month') parts[0] = part.value
      if (part.type === 'day') parts[1] = part.value
      return parts
    },
    ['', ''],
  )
  return `${dayPart} ${monthPart}`
}
