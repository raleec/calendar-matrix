import { useState } from 'react'
import { Legend } from './components/Legend'
import { Toolbar } from './components/Toolbar'
import { MatrixGrid } from './components/MatrixGrid'

function App() {
  const now = new Date()
  const [month, setMonth] = useState(now.getMonth())
  const [year, setYear] = useState(now.getFullYear())

  return (
    <main className="app">
      <header className="app-header">
        <h1>Calendar Matrix</h1>
        <Legend />
      </header>

      <Toolbar
        month={month}
        year={year}
        onMonthChange={setMonth}
        onYearChange={setYear}
      />

      <MatrixGrid month={month} year={year} />
    </main>
  )
}

export default App
