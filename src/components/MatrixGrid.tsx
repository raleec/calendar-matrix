import type { GraphPerson } from '../graph/types'
import { daysInMonth } from '../utils/date'

export interface MatrixGridProps {
  month: number
  year: number
  people?: GraphPerson[]
}

export function MatrixGrid({ month, year, people = [] }: MatrixGridProps) {
  const days = Array.from(
    { length: daysInMonth(year, month) },
    (_, index) => index + 1,
  )

  return (
    <div className="grid-area">
      <table className="matrix">
        <thead>
          <tr>
            <th scope="col" className="person-col">
              Person
            </th>
            {days.map((day) => (
              <th key={day} scope="col">
                {day}
              </th>
            ))}
            <th scope="col">Total</th>
          </tr>
        </thead>
        <tbody>
          {people.length === 0 ? (
            <tr>
              <td className="empty-state" colSpan={days.length + 2}>
                Add people to get started
              </td>
            </tr>
          ) : (
            people.map((person) => (
              <tr key={person.id}>
                <th scope="row" className="person-col">
                  {person.displayName}
                </th>
                {days.map((day) => (
                  <td key={day} />
                ))}
                <td />
              </tr>
            ))
          )}
        </tbody>
      </table>
    </div>
  )
}
