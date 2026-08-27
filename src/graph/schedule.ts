/**
 * Fetches and refines month-by-month availability data from Microsoft Graph
 * into the per-day status values rendered by the calendar matrix.
 */
import { Office365UsersService } from '../generated/services/Office365UsersService'
import type { StatusKey } from '../status'
import { daysInMonth } from '../utils/date'
import {
  getCachedSchedule,
  invalidateCachedSchedule,
  scheduleCacheKey,
  setCachedSchedule,
} from './scheduleCache'

/** Graph's `getSchedule` accepts at most this many mailboxes per request. */
const SCHEDULES_PER_REQUEST = 20
/** Number of `getSchedule` requests to keep in flight at once. */
const DEFAULT_CONCURRENCY = 4

/** A day's status once free/blank days have been resolved to `null`. */
export type DayStatus = StatusKey | null

/** Sentinel used when a person's availability could not be retrieved. */
export const UNAVAILABLE = 'unavailable' as const

/** The status shown in a single matrix cell. */
export type CellStatus = DayStatus | typeof UNAVAILABLE

/** Schedule data for a single person across the requested month. */
export interface PersonSchedule {
  /** The mailbox identifier (email address) Graph reported this schedule for. */
  scheduleId: string
  /** One entry per day of the requested month, in calendar order (index 0 = the 1st). */
  days: CellStatus[]
  /**
   * Human-readable reason the schedule could not be retrieved, shown as a
   * tooltip on {@link UNAVAILABLE} cells (e.g. "No permission to view this
   * calendar"). Unset when every day resolved successfully.
   */
  unavailableReason?: string
}

/** The minimum person identity data needed to request schedule information. */
export interface PersonInput {
  id: string
  mail: string | null
}

/** Optional knobs for how Graph schedule data is fetched and interpreted. */
export interface GetSchedulesOptions {
  /**
   * Status to use for `tentative` (availabilityView code `1`) days. Defaults
   * to `null` (blank), matching the default treatment of `free` days.
   */
  tentativeStatus?: CellStatus
  /** Maximum number of `getSchedule` requests to issue concurrently. */
  concurrency?: number
}

interface GraphScheduleItem {
  status?: string
  subject?: string
  isAllDay?: boolean
  start?: { dateTime?: string }
  end?: { dateTime?: string }
}

interface GraphScheduleInformation {
  scheduleId: string
  availabilityView?: string
  scheduleItems?: GraphScheduleItem[]
  error?: { message?: string } | null
}


function codeToStatus(code: string, tentativeStatus: CellStatus): CellStatus {
  switch (code) {
    case '1': // tentative
      return tentativeStatus
    case '4': // workingElsewhere
      return 'WE'
    case '0': // free
    case '2': // busy — treated as free
    case '3': // oof — refined via scheduleItems below
    default:
      return null
  }
}

function pad2(value: number): string {
  return String(value).padStart(2, '0')
}

/** Formats a local (Prefer: outlook.timezone="UTC") date-time for the Graph request body. */
function formatDateTime(year: number, month: number, day: number): string {
  return `${year}-${pad2(month + 1)}-${pad2(day)}T00:00:00`
}

function buildTimeWindow(
  year: number,
  month: number,
): {
  startTime: { dateTime: string; timeZone: string }
  endTime: { dateTime: string; timeZone: string }
} {
  const endMonth = month === 11 ? 0 : month + 1
  const endYear = month === 11 ? year + 1 : year

  return {
    startTime: { dateTime: formatDateTime(year, month, 1), timeZone: 'UTC' },
    endTime: {
      dateTime: formatDateTime(endYear, endMonth, 1),
      timeZone: 'UTC',
    },
  }
}

/** Default tooltip shown for an unavailable schedule when no reason (or a generic one) applies. */
export const NO_PERMISSION_REASON = 'No permission to view this calendar'

const TRAVEL_PATTERN = /\btravel\b|trip|onsite|offsite/i
const VACATION_PATTERN = /vacation|pto|annual leave|dto|out of office/i

/** Refines an OOF calendar event to 'TR' or 'DTO' based on subject, defaulting to 'OOO'. */
function oofSubjectToStatus(subject: string | undefined): 'TR' | 'DTO' | 'OOO' {
  if (!subject) return 'OOO'
  if (TRAVEL_PATTERN.test(subject)) return 'TR'
  if (VACATION_PATTERN.test(subject)) return 'DTO'
  return 'OOO'
}

/** Parses the `YYYY-MM-DD` prefix of a Graph date-time string into a day-of-month index. */
function dayIndexInMonth(dateTime: string, year: number, month: number): number {
  const [y, m, d] = dateTime.slice(0, 10).split('-').map(Number)
  return Math.round(
    (Date.UTC(y, m - 1, d) - Date.UTC(year, month, 1)) / 86_400_000,
  )
}

function unavailableSchedule(
  scheduleId: string,
  count: number,
  reason: string = NO_PERMISSION_REASON,
): PersonSchedule {
  return {
    scheduleId,
    days: Array<CellStatus>(count).fill(UNAVAILABLE),
    unavailableReason: reason,
  }
}

function toPersonSchedule(
  entry: GraphScheduleInformation,
  year: number,
  month: number,
  count: number,
  tentativeStatus: CellStatus,
): PersonSchedule {
  if (entry.error) {
    return unavailableSchedule(entry.scheduleId, count, entry.error.message)
  }

  const view = entry.availabilityView ?? ''
  const days: CellStatus[] = []
  for (let i = 0; i < count; i++) {
    days.push(codeToStatus(view[i] ?? '0', tentativeStatus))
  }

  // Refine OOF days and set WE from schedule items.
  for (const item of entry.scheduleItems ?? []) {
    const startDateTime = item.start?.dateTime
    if (!startDateTime) continue
    const startIdx = dayIndexInMonth(startDateTime, year, month)
    const endIdx = item.end?.dateTime
      ? dayIndexInMonth(item.end.dateTime, year, month)
      : startIdx + 1
    const from = Math.max(0, startIdx)
    const to = Math.min(count, Math.max(endIdx, startIdx + 1))

    let status: CellStatus
    if (item.status === 'oof') {
      status = oofSubjectToStatus(item.subject)
    } else if (item.status === 'workingElsewhere') {
      status = 'WE'
    } else {
      continue
    }

    for (let i = from; i < to; i++) {
      // TR/DTO take priority over OOO/WE
      if (days[i] === null || days[i] === 'WE' || days[i] === 'OOO') {
        days[i] = status
      }
    }
  }

  return { scheduleId: entry.scheduleId, days }
}

function chunk<T>(items: T[], size: number): T[][] {
  const chunks: T[][] = []
  for (let i = 0; i < items.length; i += size) {
    chunks.push(items.slice(i, i + size))
  }
  return chunks
}

/** Runs `worker` over `items` with at most `limit` concurrent invocations. */
async function runWithConcurrency<T>(
  items: T[],
  limit: number,
  worker: (item: T) => Promise<void>,
): Promise<void> {
  let index = 0

  async function runNext(): Promise<void> {
    const current = index++
    if (current >= items.length) return
    await worker(items[current])
    await runNext()
  }

  const workerCount = Math.max(1, Math.min(limit, items.length))
  await Promise.all(Array.from({ length: workerCount }, () => runNext()))
}

/**
 * Calls `/me/calendar/getSchedule` for a single chunk of mailboxes via the
 * Office 365 Users connector's HttpRequest operation.
 */
async function postGetSchedule(
  schedules: string[],
  timeWindow: ReturnType<typeof buildTimeWindow>,
): Promise<GraphScheduleInformation[]> {
  // Body must be a plain object (not a JSON string) — the connector schema
  // declares Body as type "object" and the SDK serializes it to JSON internally.
  const bodyObj = {
    schedules,
    startTime: timeWindow.startTime,
    endTime: timeWindow.endTime,
    availabilityViewInterval: 1440, // minutes per slot; 24h yields one availability code per day
  }

  const result = await Office365UsersService.HttpRequest(
    'https://graph.microsoft.com/v1.0/me/calendar/getSchedule',
    'POST',
    bodyObj as unknown as string,
    'application/json',
    'Prefer: outlook.timezone="UTC"',
  )

  if (!result.success) {
    const err = result.error as { message?: string; status?: number } | undefined
    throw new Error(
      `getSchedule failed${err?.status ? ` (HTTP ${err.status})` : ''}: ${err?.message ?? 'unknown error'}`,
    )
  }

  const data = result.data as { value?: GraphScheduleInformation[] }
  return data?.value ?? []
}

/**
 * Resolves per-day availability for a set of people in a given month.
 *
 * Mailboxes are chunked into groups of {@link SCHEDULES_PER_REQUEST} (the
 * Graph limit) and fetched with a small concurrency cap. Results are cached
 * in-memory and in `sessionStorage`, keyed by person/year/month, so repeat
 * requests (e.g. re-rendering the same month) avoid a network round trip.
 * Per-person failures (missing mailbox, no permission, or a request-level
 * failure) degrade to an {@link UNAVAILABLE} schedule instead of throwing.
 */
export async function getSchedules(
  people: PersonInput[],
  year: number,
  month: number,
  options: GetSchedulesOptions = {},
): Promise<Map<string, PersonSchedule>> {
  const results = new Map<string, PersonSchedule>()
  const dayCount = daysInMonth(year, month)
  const tentativeStatus = options.tentativeStatus ?? null
  const concurrency = options.concurrency ?? DEFAULT_CONCURRENCY

  const toFetch: PersonInput[] = []
  for (const person of people) {
    const cacheKey = scheduleCacheKey(person.id, year, month)
    const cached = getCachedSchedule(cacheKey)
    if (cached) {
      results.set(person.id, cached)
      continue
    }

    if (!person.mail) {
      results.set(
        person.id,
        unavailableSchedule('', dayCount, 'No email address on file'),
      )
      continue
    }

    toFetch.push(person)
  }

  if (toFetch.length === 0) return results

  const timeWindow = buildTimeWindow(year, month)
  const chunks = chunk(toFetch, SCHEDULES_PER_REQUEST)

  let firstError: string | null = null

  await runWithConcurrency(chunks, concurrency, async (peopleChunk) => {
    const mails = peopleChunk.map((person) => person.mail as string)
    let entries: GraphScheduleInformation[]
    try {
      entries = await postGetSchedule(mails, timeWindow)
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Could not reach Microsoft Graph.'
      if (!firstError) firstError = msg
      entries = mails.map((mail) => ({
        scheduleId: mail,
        error: { message: msg },
      }))
    }

    for (let i = 0; i < peopleChunk.length; i++) {
      const person = peopleChunk[i]
      const entry = entries[i] ?? {
        scheduleId: person.mail as string,
        error: { message: 'No response received for this schedule.' },
      }
      const schedule = toPersonSchedule(
        entry,
        year,
        month,
        dayCount,
        tentativeStatus,
      )
      results.set(person.id, schedule)
      setCachedSchedule(scheduleCacheKey(person.id, year, month), schedule)
    }
  })

  if (firstError) {
    throw new Error(firstError)
  }

  return results
}

/**
 * Busts the cached schedule for each of `people` for the given month, so the
 * next {@link getSchedules} call re-fetches fresh data from Graph. Used by
 * the grid's "Refresh" action.
 */
export function invalidateSchedules(
  people: PersonInput[],
  year: number,
  month: number,
): void {
  for (const person of people) {
    invalidateCachedSchedule(scheduleCacheKey(person.id, year, month))
  }
}
