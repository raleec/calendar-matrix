/**
 * Shared status code definitions used throughout the calendar matrix.
 *
 * `StatusKey` is the canonical union of supported status codes, while
 * {@link LegendEntry} and {@link LEGEND} keep the user-facing labels aligned
 * with that constrained set of keys.
 */
export type StatusKey = 'OOO' | 'TR' | 'DTO' | 'WE'

/** A single legend row pairing a status code with its display label. */
export interface LegendEntry {
  key: StatusKey
  label: string
}

/** The ordered set of status codes and labels shown in the matrix legend. */
export const LEGEND: readonly LegendEntry[] = [
  { key: 'OOO', label: 'Out of Office' },
  { key: 'DTO', label: 'Vacation' },
  { key: 'TR', label: 'Travel' },
  { key: 'WE', label: 'Working Elsewhere' },
]

/** Maps a status letter code to its full legend label, e.g. `'OOO'` -> `'Out of Office'`. */
export function statusLabel(key: StatusKey): string {
  return LEGEND.find((entry) => entry.key === key)?.label ?? key
}
