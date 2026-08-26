export type StatusKey = 'V' | 'P' | 'T' | 'WE'

export interface LegendEntry {
  key: StatusKey
  label: string
}

export const LEGEND: readonly LegendEntry[] = [
  { key: 'V', label: 'Vacation' },
  { key: 'P', label: 'Personal Leave' },
  { key: 'T', label: 'Travel' },
  { key: 'WE', label: 'Working Elsewhere' },
]

/** Maps a status letter code to its full legend label, e.g. `'V'` -> `'Vacation'`. */
export function statusLabel(key: StatusKey): string {
  return LEGEND.find((entry) => entry.key === key)?.label ?? key
}
