import { useState } from 'react'
import { Legend } from './components/Legend'
import { Toolbar } from './components/Toolbar'
import { MatrixGrid } from './components/MatrixGrid'
import { AccountButton } from './components/AccountButton'
import { PeoplePicker } from './components/PeoplePicker'
import { usePeopleSelection } from './hooks/usePeopleSelection'

function App() {
  const now = new Date()
  const [month, setMonth] = useState(now.getMonth())
  const [year, setYear] = useState(now.getFullYear())
  const peopleSelection = usePeopleSelection()

  return (
    <main className="app">
      <header className="app-header">
        <h1>Calendar Matrix</h1>
        <AccountButton />
        <Legend />
      </header>

      <Toolbar
        month={month}
        year={year}
        onMonthChange={setMonth}
        onYearChange={setYear}
      />

      <PeoplePicker selection={peopleSelection} />

      <MatrixGrid month={month} year={year} people={peopleSelection.people} />
    </main>
  )
}

export default App
