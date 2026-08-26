import type { PersonSchedule } from './schedule'

const STORAGE_PREFIX = 'calendar-matrix.schedule'

/** Fast in-memory layer in front of `sessionStorage` for the current tab. */
const memoryCache = new Map<string, PersonSchedule>()

/** Cache key for a person's schedule in a given calendar month. */
export function scheduleCacheKey(
  personId: string,
  year: number,
  month: number,
): string {
  return `${personId}:${year}:${month}`
}

function storageKey(key: string): string {
  return `${STORAGE_PREFIX}:${key}`
}

/**
 * Reads a cached schedule, checking the in-memory cache first and falling
 * back to `sessionStorage` (populating the in-memory cache on hit).
 */
export function getCachedSchedule(key: string): PersonSchedule | undefined {
  const cached = memoryCache.get(key)
  if (cached) return cached

  try {
    const raw = sessionStorage.getItem(storageKey(key))
    if (!raw) return undefined
    const parsed = JSON.parse(raw) as PersonSchedule
    memoryCache.set(key, parsed)
    return parsed
  } catch {
    // sessionStorage may be unavailable (private browsing, SSR) or contain
    // corrupt data — treat as a cache miss either way.
    return undefined
  }
}

/** Writes a schedule to both the in-memory cache and `sessionStorage`. */
export function setCachedSchedule(key: string, schedule: PersonSchedule): void {
  memoryCache.set(key, schedule)
  try {
    sessionStorage.setItem(storageKey(key), JSON.stringify(schedule))
  } catch {
    // Quota exceeded or storage disabled — the in-memory cache still works
    // for the lifetime of this page.
  }
}

/** Clears the in-memory cache. Primarily useful for tests. */
export function clearScheduleCache(): void {
  memoryCache.clear()
}
