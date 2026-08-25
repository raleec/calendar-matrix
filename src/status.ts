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
