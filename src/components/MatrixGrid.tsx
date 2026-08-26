import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { useMsal } from '@azure/msal-react'
import type { GraphPerson } from '../graph/types'
import { createGraphClient } from '../graph/client'
import { useGraphToken } from '../hooks/useGraphToken'
import {
  getSchedules,
  invalidateSchedules,
  UNAVAILABLE,
  type CellStatus,
  type PersonSchedule,
} from '../graph/schedule'
import { daysInMonth, formatDayLabel } from '../utils/date'
import { statusLabel } from '../status'
import { buildMatrixCsv, downloadCsv } from '../utils/csv'

export interface MatrixGridProps {
  month: number
  year: number
  people?: GraphPerson[]
}

const WEEKDAY_FORMATTER = new Intl.DateTimeFormat('en-US', {
  weekday: 'short',
})

const MONTH_FORMATTER = new Intl.DateTimeFormat('en-US', {
  month: 'long',
  year: 'numeric',
})

function isWeekend(year: number, month: number, day: number): boolean {
  const weekday = new Date(year, month, day).getDay()
  return weekday === 0 || weekday === 6
}

function isAbsence(status: CellStatus | undefined): boolean {
  return status === 'V' || status === 'P' || status === 'T' || status === 'WE'
}

/** Builds the `aria-label` announced for a single grid cell. */
function cellAriaLabel(
  personName: string,
  dayLabel: string,
  status: CellStatus | undefined,
  unavailableReason: string | undefined,
): string {
  if (!status) return `${personName}, ${dayLabel}`
  if (status === UNAVAILABLE) {
    return `${personName}, ${dayLabel}, ${unavailableReason ?? 'Unavailable'}`
  }
  return `${personName}, ${dayLabel}, ${statusLabel(status)}`
}

export function MatrixGrid({ month, year, people = [] }: MatrixGridProps) {
  const { accounts } = useMsal()
  const getGraphToken = useGraphToken()
  const graphClient = useMemo(
    () => createGraphClient(() => getGraphToken()),
    [getGraphToken],
  )

  const [schedules, setSchedules] = useState<Map<string, PersonSchedule>>(
    new Map(),
  )
  const [isLoading, setIsLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [reloadToken, setReloadToken] = useState(0)

  const dayCount = daysInMonth(year, month)
  const days = useMemo(
    () => Array.from({ length: dayCount }, (_, index) => index + 1),
    [dayCount],
  )

  const cellRefs = useRef<(HTMLTableCellElement | null)[][]>([])

  useEffect(() => {
    if (people.length === 0 || !accounts[0]) {
      return
    }

    let cancelled = false
    const controller = new AbortController()

    async function loadSchedules() {
      setIsLoading(true)
      setError(null)
      try {
        const result = await getSchedules(
          graphClient,
          people.map((person) => ({ id: person.id, mail: person.mail })),
          year,
          month,
          { signal: controller.signal },
        )
        if (!cancelled) setSchedules(result)
      } catch (err) {
        if (controller.signal.aborted) return
        if (!cancelled) {
          setSchedules(new Map())
          setError(
            err instanceof Error
              ? err.message
              : 'Failed to load calendars from Microsoft Graph.',
          )
        }
      } finally {
        if (!cancelled) setIsLoading(false)
      }
    }

    void loadSchedules()

    return () => {
      cancelled = true
      controller.abort()
    }
  }, [graphClient, people, year, month, accounts, reloadToken])

  const handleRetry = useCallback(() => {
    setReloadToken((token) => token + 1)
  }, [])

  const handleRefresh = useCallback(() => {
    invalidateSchedules(
      people.map((person) => ({ id: person.id, mail: person.mail })),
      year,
      month,
    )
    setReloadToken((token) => token + 1)
  }, [people, year, month])

  const rowTotals = useMemo(() => {
    const totals = new Map<string, number>()
    for (const person of people) {
      const schedule = schedules.get(person.id)
      const count = schedule
        ? schedule.days.filter((status) => isAbsence(status)).length
        : 0
      totals.set(person.id, count)
    }
    return totals
  }, [people, schedules])

  const columnTotals = useMemo(
    () =>
      days.map((day) =>
        people.reduce((count, person) => {
          const status = schedules.get(person.id)?.days[day - 1]
          return isAbsence(status) ? count + 1 : count
        }, 0),
      ),
    [days, people, schedules],
  )

  const grandTotal = columnTotals.reduce((sum, value) => sum + value, 0)
  const columnCount = days.length + 2
  const showTotals = people.length > 0 && !isLoading

  const handleExportCsv = useCallback(() => {
    const csv = buildMatrixCsv(
      people,
      days,
      schedules,
      rowTotals,
      columnTotals,
      grandTotal,
    )
    const monthLabel = MONTH_FORMATTER.format(new Date(year, month, 1)).replace(
      ' ',
      '-',
    )
    downloadCsv(`calendar-matrix-${monthLabel}.csv`, csv)
  }, [
    people,
    days,
    schedules,
    rowTotals,
    columnTotals,
    grandTotal,
    year,
    month,
  ])

  const focusCell = (rowIndex: number, colIndex: number) => {
    const row = cellRefs.current[rowIndex]
    const cell = row?.[Math.max(0, Math.min(colIndex, days.length - 1))]
    cell?.focus()
  }

  const handleCellKeyDown = (
    event: React.KeyboardEvent<HTMLTableCellElement>,
    rowIndex: number,
    colIndex: number,
  ) => {
    switch (event.key) {
      case 'ArrowRight':
        event.preventDefault()
        focusCell(rowIndex, colIndex + 1)
        break
      case 'ArrowLeft':
        event.preventDefault()
        focusCell(rowIndex, colIndex - 1)
        break
      case 'ArrowDown':
        event.preventDefault()
        focusCell(Math.min(rowIndex + 1, people.length - 1), colIndex)
        break
      case 'ArrowUp':
        event.preventDefault()
        focusCell(Math.max(rowIndex - 1, 0), colIndex)
        break
      case 'Home':
        event.preventDefault()
        focusCell(rowIndex, 0)
        break
      case 'End':
        event.preventDefault()
        focusCell(rowIndex, days.length - 1)
        break
      default:
        break
    }
  }

  return (
    <div className="grid-area-wrapper">
      <div className="grid-toolbar">
        <button
          type="button"
          onClick={handleRefresh}
          disabled={people.length === 0 || isLoading}
        >
          Refresh
        </button>
        <button type="button" onClick={handleExportCsv} disabled={!showTotals}>
          Export CSV
        </button>
      </div>

      {error && people.length > 0 && (
        <div className="error-banner" role="alert">
          <span>Couldn&apos;t load calendars: {error}</span>
          <button type="button" onClick={handleRetry}>
            Retry
          </button>
        </div>
      )}

      <div className="grid-area">
        <table className="matrix">
          <thead>
            <tr>
              <th scope="col" className="person-col sticky-col" rowSpan={2}>
                Person
              </th>
              {days.map((day) => (
                <th
                  key={`weekday-${day}`}
                  scope="col"
                  className={
                    isWeekend(year, month, day) ? 'weekend' : undefined
                  }
                >
                  {WEEKDAY_FORMATTER.format(new Date(year, month, day))}
                </th>
              ))}
              <th scope="col" className="total-col" rowSpan={2}>
                Total
              </th>
            </tr>
            <tr>
              {days.map((day) => (
                <th
                  key={day}
                  scope="col"
                  className={
                    isWeekend(year, month, day) ? 'weekend' : undefined
                  }
                >
                  {day}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {people.length === 0 ? (
              <tr>
                <td className="empty-state" colSpan={columnCount}>
                  Add people to get started
                </td>
              </tr>
            ) : isLoading ? (
              people.map((person) => (
                <tr key={`skeleton-${person.id}`} className="skeleton-row">
                  <th scope="row" className="person-col sticky-col">
                    {person.displayName}
                  </th>
                  {days.map((day) => (
                    <td
                      key={day}
                      className={
                        isWeekend(year, month, day) ? 'weekend' : undefined
                      }
                    >
                      <span className="skeleton-cell" aria-hidden="true" />
                    </td>
                  ))}
                  <td className="total-col">
                    <span className="skeleton-cell" aria-hidden="true" />
                  </td>
                </tr>
              ))
            ) : (
              people.map((person, rowIndex) => {
                const schedule = schedules.get(person.id)
                return (
                  <tr key={person.id}>
                    <th scope="row" className="person-col sticky-col">
                      {person.displayName}
                    </th>
                    {days.map((day, colIndex) => {
                      const status = schedule?.days[day - 1]
                      const dayLabel = formatDayLabel(year, month, day)
                      const classNames = [
                        isWeekend(year, month, day) ? 'weekend' : '',
                        status && status !== UNAVAILABLE
                          ? `cell-${status}`
                          : '',
                        status === UNAVAILABLE ? 'cell-unavailable' : '',
                      ]
                        .filter(Boolean)
                        .join(' ')
                      return (
                        <td
                          key={day}
                          ref={(node) => {
                            cellRefs.current[rowIndex] ??= []
                            cellRefs.current[rowIndex][colIndex] = node
                          }}
                          className={classNames || undefined}
                          tabIndex={rowIndex === 0 && colIndex === 0 ? 0 : -1}
                          aria-label={cellAriaLabel(
                            person.displayName,
                            dayLabel,
                            status,
                            schedule?.unavailableReason,
                          )}
                          title={
                            status === UNAVAILABLE
                              ? (schedule?.unavailableReason ?? 'Unavailable')
                              : undefined
                          }
                          onKeyDown={(event) =>
                            handleCellKeyDown(event, rowIndex, colIndex)
                          }
                        >
                          {status && status !== UNAVAILABLE ? status : ''}
                        </td>
                      )
                    })}
                    <td className="total-col">
                      {rowTotals.get(person.id) ?? 0}
                    </td>
                  </tr>
                )
              })
            )}
          </tbody>
          {showTotals && (
            <tfoot>
              <tr>
                <th scope="row" className="person-col sticky-col">
                  Total
                </th>
                {columnTotals.map((total, index) => (
                  <td
                    key={days[index]}
                    className={
                      isWeekend(year, month, days[index])
                        ? 'weekend'
                        : undefined
                    }
                  >
                    {total}
                  </td>
                ))}
                <td className="total-col">{grandTotal}</td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </div>
  )
}
