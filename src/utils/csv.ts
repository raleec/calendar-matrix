import type { GraphPerson } from '../graph/types'
import {
  UNAVAILABLE,
  type CellStatus,
  type PersonSchedule,
} from '../graph/schedule'

/** Escapes a single CSV field per RFC 4180 (quoting only when necessary). */
function csvField(value: string | number): string {
  const text = String(value)
  if (/[",\r\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`
  }
  return text
}

function cellText(status: CellStatus | undefined): string {
  if (!status || status === UNAVAILABLE) return ''
  return status
}

/**
 * Builds a CSV (RFC 4180, CRLF line endings) of the current grid: one row
 * per person with a column per day of the month plus a trailing total, and a
 * final "Total" row with per-day and grand totals. Opens cleanly in Excel.
 */
export function buildMatrixCsv(
  people: GraphPerson[],
  days: number[],
  schedules: Map<string, PersonSchedule>,
  rowTotals: Map<string, number>,
  columnTotals: number[],
  grandTotal: number,
): string {
  const rows: string[][] = []

  rows.push(['Person', ...days.map((day) => String(day)), 'Total'])

  for (const person of people) {
    const schedule = schedules.get(person.id)
    rows.push([
      person.displayName,
      ...days.map((day) => cellText(schedule?.days[day - 1])),
      String(rowTotals.get(person.id) ?? 0),
    ])
  }

  if (people.length > 0) {
    rows.push([
      'Total',
      ...columnTotals.map((total) => String(total)),
      String(grandTotal),
    ])
  }

  return rows.map((row) => row.map(csvField).join(',')).join('\r\n')
}

/** Triggers a browser download of `content` as a file named `filename`. */
export function downloadCsv(filename: string, content: string): void {
  // Prepend a UTF-8 BOM so Excel detects the encoding correctly.
  const blob = new Blob(['\uFEFF', content], {
    type: 'text/csv;charset=utf-8;',
  })
  const url = URL.createObjectURL(blob)
  try {
    const link = document.createElement('a')
    link.href = url
    link.download = filename
    document.body.appendChild(link)
    link.click()
    link.remove()
  } finally {
    URL.revokeObjectURL(url)
  }
}
