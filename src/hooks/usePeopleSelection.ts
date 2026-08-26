import { useCallback, useEffect, useState } from 'react'
import type { GraphPerson } from '../graph/types'

const STORAGE_KEY = 'calendar-matrix.selectedPeople'

function loadStoredSelection(): GraphPerson[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return []
    const parsed = JSON.parse(raw)
    if (!Array.isArray(parsed)) return []
    return parsed.filter(
      (item): item is GraphPerson =>
        item &&
        typeof item.id === 'string' &&
        typeof item.displayName === 'string',
    )
  } catch {
    return []
  }
}

/**
 * Manages the set of people whose availability is shown in the matrix.
 * Selections are de-duplicated by user id and persisted to `localStorage`
 * so they survive a page reload.
 */
export function usePeopleSelection() {
  const [people, setPeople] = useState<GraphPerson[]>(() =>
    loadStoredSelection(),
  )

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(people))
  }, [people])

  const addPeople = useCallback((toAdd: GraphPerson[]) => {
    if (toAdd.length === 0) return
    setPeople((current) => {
      const byId = new Map(current.map((person) => [person.id, person]))
      for (const person of toAdd) {
        byId.set(person.id, person)
      }
      return Array.from(byId.values())
    })
  }, [])

  const removePerson = useCallback((id: string) => {
    setPeople((current) => current.filter((person) => person.id !== id))
  }, [])

  const clearAll = useCallback(() => {
    setPeople([])
  }, [])

  return { people, addPeople, removePerson, clearAll }
}
