import { useState } from 'react'
import {
  AuthenticatedTemplate,
  UnauthenticatedTemplate,
} from '@azure/msal-react'
import { Legend } from './components/Legend'
import { Toolbar } from './components/Toolbar'
import { MatrixGrid } from './components/MatrixGrid'
import { AccountButton } from './components/AccountButton'

function App() {
  const now = new Date()
  const [month, setMonth] = useState(now.getMonth())
  const [year, setYear] = useState(now.getFullYear())

  return (
    <main className="app">
      <header className="app-header">
        <h1>Calendar Matrix</h1>
        <AccountButton />
        <Legend />
      </header>

      <AuthenticatedTemplate>
        <Toolbar
          month={month}
          year={year}
          onMonthChange={setMonth}
          onYearChange={setYear}
        />

        <MatrixGrid month={month} year={year} />
      </AuthenticatedTemplate>

      <UnauthenticatedTemplate>
        <p className="signed-out-message">
          Sign in with your corporate account to view the calendar matrix.
        </p>
      </UnauthenticatedTemplate>
    </main>
  )
}

export default App
