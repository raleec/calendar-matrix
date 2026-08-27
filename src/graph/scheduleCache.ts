import type { PersonSchedule } from './schedule'

const STORAGE_PREFIX = 'calendar-matrix.schedule'

/** Cached entries expire after this many milliseconds (30 minutes). */
const CACHE_TTL_MS = 30 * 60 * 1000

interface CacheEntry {
  schedule: PersonSchedule
  cachedAt: number
}

/** Fast in-memory layer in front of `sessionStorage` for the current tab. */
const memoryCache = new Map<string, CacheEntry>()

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

function isFresh(entry: CacheEntry): boolean {
  return Date.now() - entry.cachedAt < CACHE_TTL_MS
}

/**
 * Reads a cached schedule, checking the in-memory cache first and falling
 * back to `sessionStorage` (populating the in-memory cache on hit).
 * Returns `undefined` if the entry is missing or older than {@link CACHE_TTL_MS}.
 */
export function getCachedSchedule(key: string): PersonSchedule | undefined {
  const mem = memoryCache.get(key)
  if (mem) {
    if (isFresh(mem)) return mem.schedule
    memoryCache.delete(key)
  }

  try {
    const raw = sessionStorage.getItem(storageKey(key))
    if (!raw) return undefined
    const entry = JSON.parse(raw) as CacheEntry
    if (!isFresh(entry)) {
      sessionStorage.removeItem(storageKey(key))
      return undefined
    }
    memoryCache.set(key, entry)
    return entry.schedule
  } catch {
    // sessionStorage may be unavailable (private browsing, SSR) or contain
    // corrupt data — treat as a cache miss either way.
    return undefined
  }
}

/** Writes a schedule to both the in-memory cache and `sessionStorage`. */
export function setCachedSchedule(key: string, schedule: PersonSchedule): void {
  const entry: CacheEntry = { schedule, cachedAt: Date.now() }
  memoryCache.set(key, entry)
  try {
    sessionStorage.setItem(storageKey(key), JSON.stringify(entry))
  } catch {
    // Quota exceeded or storage disabled — the in-memory cache still works
    // for the lifetime of this page.
  }
}

/** Clears the in-memory cache. Primarily useful for tests. */
export function clearScheduleCache(): void {
  memoryCache.clear()
}

/**
 * Removes the cached schedule (both in-memory and `sessionStorage`) for a
 * single person/month, so the next `getSchedules` call re-fetches it from
 * Graph. Used by the grid's refresh action.
 */
export function invalidateCachedSchedule(key: string): void {
  memoryCache.delete(key)
  try {
    sessionStorage.removeItem(storageKey(key))
  } catch {
    // sessionStorage may be unavailable — the in-memory cache entry is
    // already cleared, so there is nothing further to do.
  }
}
