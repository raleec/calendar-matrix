export interface MatrixGridProps {
  month: number
  year: number
}

function daysInMonth(year: number, month: number) {
  return new Date(year, month + 1, 0).getDate()
}

export function MatrixGrid({ month, year }: MatrixGridProps) {
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
          <tr>
            <td className="empty-state" colSpan={days.length + 2}>
              Add people to get started
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  )
}
