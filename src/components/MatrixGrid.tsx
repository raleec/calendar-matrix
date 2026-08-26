import { useEffect, useMemo, useState } from 'react'
import { useMsal } from '@azure/msal-react'
import type { GraphPerson } from '../graph/types'
import { createGraphClient } from '../graph/client'
import { useGraphToken } from '../hooks/useGraphToken'
import {
  getSchedules,
  UNAVAILABLE,
  type CellStatus,
  type PersonSchedule,
} from '../graph/schedule'
import { daysInMonth } from '../utils/date'

export interface MatrixGridProps {
  month: number
  year: number
  people?: GraphPerson[]
}

const WEEKDAY_FORMATTER = new Intl.DateTimeFormat('en-US', {
  weekday: 'short',
})

function isWeekend(year: number, month: number, day: number): boolean {
  const weekday = new Date(year, month, day).getDay()
  return weekday === 0 || weekday === 6
}

function isAbsence(status: CellStatus | undefined): boolean {
  return status === 'V' || status === 'P' || status === 'T' || status === 'WE'
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

  const dayCount = daysInMonth(year, month)
  const days = useMemo(
    () => Array.from({ length: dayCount }, (_, index) => index + 1),
    [dayCount],
  )

  useEffect(() => {
    if (people.length === 0 || !accounts[0]) {
      return
    }

    let cancelled = false
    const controller = new AbortController()

    async function loadSchedules() {
      setIsLoading(true)
      try {
        const result = await getSchedules(
          graphClient,
          people.map((person) => ({ id: person.id, mail: person.mail })),
          year,
          month,
          { signal: controller.signal },
        )
        if (!cancelled) setSchedules(result)
      } catch {
        if (!cancelled) setSchedules(new Map())
      } finally {
        if (!cancelled) setIsLoading(false)
      }
    }

    void loadSchedules()

    return () => {
      cancelled = true
      controller.abort()
    }
  }, [graphClient, people, year, month, accounts])

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

  return (
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
                className={isWeekend(year, month, day) ? 'weekend' : undefined}
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
                className={isWeekend(year, month, day) ? 'weekend' : undefined}
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
            people.map((person) => {
              const schedule = schedules.get(person.id)
              return (
                <tr key={person.id}>
                  <th scope="row" className="person-col sticky-col">
                    {person.displayName}
                  </th>
                  {days.map((day) => {
                    const status = schedule?.days[day - 1]
                    const classNames = [
                      isWeekend(year, month, day) ? 'weekend' : '',
                      status && status !== UNAVAILABLE ? `cell-${status}` : '',
                      status === UNAVAILABLE ? 'cell-unavailable' : '',
                    ]
                      .filter(Boolean)
                      .join(' ')
                    return (
                      <td key={day} className={classNames || undefined}>
                        {status && status !== UNAVAILABLE ? status : ''}
                      </td>
                    )
                  })}
                  <td className="total-col">{rowTotals.get(person.id) ?? 0}</td>
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
                    isWeekend(year, month, days[index]) ? 'weekend' : undefined
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
  )
}
