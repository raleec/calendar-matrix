import { useEffect, useMemo, useRef, useState } from 'react'
import { useMsal } from '@azure/msal-react'
import { createGraphClient } from '../graph/client'
import {
  getDirectReports,
  getGroupMembers,
  searchGroups,
  searchUsers,
} from '../graph/people'
import type { GraphGroup, GraphPerson } from '../graph/types'
import { useGraphToken } from '../hooks/useGraphToken'
import { usePeopleSelection } from '../hooks/usePeopleSelection'

const DEBOUNCE_MS = 300

type SearchResult =
  { kind: 'person'; person: GraphPerson } | { kind: 'group'; group: GraphGroup }

export interface PeoplePickerProps {
  selection: ReturnType<typeof usePeopleSelection>
}

export function PeoplePicker({ selection }: PeoplePickerProps) {
  const { people, addPeople, removePerson } = selection
  const { instance, accounts } = useMsal()
  const getGraphToken = useGraphToken()
  const graphClient = useMemo(
    () => createGraphClient(() => getGraphToken()),
    [getGraphToken],
  )

  const [query, setQuery] = useState('')
  const [results, setResults] = useState<SearchResult[]>([])
  const [isSearching, setIsSearching] = useState(false)
  const [isExpandingGroup, setIsExpandingGroup] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const abortRef = useRef<AbortController | null>(null)

  useEffect(() => {
    const term = query.trim()
    abortRef.current?.abort()

    if (!term) {
      return
    }

    const controller = new AbortController()
    abortRef.current = controller

    const timeoutId = window.setTimeout(() => {
      setIsSearching(true)
      setError(null)

      Promise.all([
        searchUsers(graphClient, term, controller.signal),
        searchGroups(graphClient, term, controller.signal),
      ])
        .then(([users, groups]) => {
          if (controller.signal.aborted) return
          setResults([
            ...groups.map((group): SearchResult => ({ kind: 'group', group })),
            ...users.map((person): SearchResult => ({
              kind: 'person',
              person,
            })),
          ])
        })
        .catch((err) => {
          if (controller.signal.aborted) return
          setError(err instanceof Error ? err.message : 'Search failed.')
        })
        .finally(() => {
          if (controller.signal.aborted) return
          setIsSearching(false)
        })
    }, DEBOUNCE_MS)

    return () => {
      window.clearTimeout(timeoutId)
      controller.abort()
    }
  }, [query, graphClient])

  const handleSelectPerson = (person: GraphPerson) => {
    addPeople([person])
    setQuery('')
    setResults([])
  }

  const handleSelectGroup = async (group: GraphGroup) => {
    setIsExpandingGroup(true)
    setError(null)
    try {
      const members = await getGroupMembers(graphClient, group.id)
      addPeople(members)
      setQuery('')
      setResults([])
    } catch (err) {
      setError(
        err instanceof Error ? err.message : 'Failed to expand group members.',
      )
    } finally {
      setIsExpandingGroup(false)
    }
  }

  const handleAddDirectReports = async () => {
    setIsExpandingGroup(true)
    setError(null)
    try {
      const reports = await getDirectReports(graphClient)
      addPeople(reports)
    } catch (err) {
      setError(
        err instanceof Error ? err.message : 'Failed to load direct reports.',
      )
    } finally {
      setIsExpandingGroup(false)
    }
  }

  const busy = isSearching || isExpandingGroup
  const activeAccount = accounts[0]

  return (
    <div className="people-picker">
      <div className="people-picker-search">
        <label className="field people-picker-field">
          People or groups
          <input
            type="text"
            value={query}
            placeholder="Search by name or email…"
            onChange={(event) => setQuery(event.target.value)}
            disabled={!instance || !activeAccount}
          />
        </label>
        <button
          type="button"
          className="people-picker-shortcut"
          onClick={() => void handleAddDirectReports()}
          disabled={busy || !activeAccount}
        >
          Add my direct reports
        </button>
      </div>

      {error && <p className="people-picker-error">{error}</p>}

      {query.trim() && (
        <ul className="people-picker-results">
          {busy && results.length === 0 && (
            <li className="people-picker-status">Searching…</li>
          )}
          {!busy && results.length === 0 && (
            <li className="people-picker-status">No matches found.</li>
          )}
          {results.map((result) =>
            result.kind === 'group' ? (
              <li key={`group-${result.group.id}`}>
                <button
                  type="button"
                  className="people-picker-result people-picker-result-group"
                  onClick={() => void handleSelectGroup(result.group)}
                  disabled={isExpandingGroup}
                >
                  <span className="people-picker-result-name">
                    {result.group.displayName}
                  </span>
                  <span className="people-picker-result-tag">Group</span>
                </button>
              </li>
            ) : (
              <li key={`person-${result.person.id}`}>
                <button
                  type="button"
                  className="people-picker-result"
                  onClick={() => handleSelectPerson(result.person)}
                >
                  <span className="people-picker-result-name">
                    {result.person.displayName}
                  </span>
                  {result.person.mail && (
                    <span className="people-picker-result-mail">
                      {result.person.mail}
                    </span>
                  )}
                </button>
              </li>
            ),
          )}
        </ul>
      )}

      {people.length > 0 && (
        <ul className="people-chip-list" aria-label="Selected people">
          {people.map((person) => (
            <li key={person.id} className="people-chip">
              <span className="people-chip-name">{person.displayName}</span>
              <button
                type="button"
                className="people-chip-remove"
                aria-label={`Remove ${person.displayName}`}
                onClick={() => removePerson(person.id)}
              >
                ×
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}
