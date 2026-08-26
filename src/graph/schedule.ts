import type { Client } from '@microsoft/microsoft-graph-client'
import type { StatusKey } from '../status'
import { daysInMonth } from '../utils/date'
import {
  getCachedSchedule,
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

export interface PersonSchedule {
  /** The mailbox identifier (email address) Graph reported this schedule for. */
  scheduleId: string
  /** One entry per day of the requested month, in calendar order (index 0 = the 1st). */
  days: CellStatus[]
}

export interface PersonInput {
  id: string
  mail: string | null
}

export interface GetSchedulesOptions {
  /**
   * Status to use for `tentative` (availabilityView code `1`) days. Defaults
   * to `null` (blank), matching the default treatment of `free` days.
   */
  tentativeStatus?: CellStatus
  /** Maximum number of `getSchedule` requests to issue concurrently. */
  concurrency?: number
  signal?: AbortSignal
}

interface GraphScheduleItem {
  status?: string
  subject?: string
  start?: { dateTime?: string }
  end?: { dateTime?: string }
}

interface GraphScheduleInformation {
  scheduleId: string
  availabilityView?: string
  scheduleItems?: GraphScheduleItem[]
  error?: { message?: string } | null
}

const VACATION_PATTERN = /vacation|pto|annual leave/i
const PERSONAL_PATTERN = /personal|appointment/i
const TRAVEL_PATTERN = /travel|trip|onsite|offsite/i

/** availabilityView codes for which a matching subject heuristic may override the default cell. */
const OVERRIDABLE_CODES = new Set(['1', '2', '3'])

function subjectToStatus(subject: string | undefined): StatusKey | null {
  if (!subject) return null
  if (VACATION_PATTERN.test(subject)) return 'V'
  if (PERSONAL_PATTERN.test(subject)) return 'P'
  if (TRAVEL_PATTERN.test(subject)) return 'T'
  return null
}

function codeToStatus(code: string, tentativeStatus: CellStatus): CellStatus {
  switch (code) {
    case '1': // tentative
      return tentativeStatus
    case '3': // oof
      return 'V'
    case '4': // workingElsewhere
      return 'WE'
    case '0': // free
    case '2': // busy
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

/** Parses the `YYYY-MM-DD` prefix of a Graph date-time string into a day-of-month index for `year`/`month`. */
function dayIndexInMonth(
  dateTime: string,
  year: number,
  month: number,
): number {
  const [y, m, d] = dateTime.slice(0, 10).split('-').map(Number)
  return Math.round(
    (Date.UTC(y, m - 1, d) - Date.UTC(year, month, 1)) / 86_400_000,
  )
}

function unavailableSchedule(
  scheduleId: string,
  count: number,
): PersonSchedule {
  return { scheduleId, days: Array<CellStatus>(count).fill(UNAVAILABLE) }
}

function toPersonSchedule(
  entry: GraphScheduleInformation,
  year: number,
  month: number,
  count: number,
  tentativeStatus: CellStatus,
): PersonSchedule {
  if (entry.error) {
    return unavailableSchedule(entry.scheduleId, count)
  }

  const view = entry.availabilityView ?? ''
  const days: CellStatus[] = []
  for (let i = 0; i < count; i++) {
    days.push(codeToStatus(view[i] ?? '0', tentativeStatus))
  }

  for (const item of entry.scheduleItems ?? []) {
    const override = subjectToStatus(item.subject)
    const startDateTime = item.start?.dateTime
    if (!override || !startDateTime) continue

    const startIdx = dayIndexInMonth(startDateTime, year, month)
    const endIdx = item.end?.dateTime
      ? dayIndexInMonth(item.end.dateTime, year, month)
      : startIdx + 1
    const from = Math.max(0, startIdx)
    const to = Math.min(count, Math.max(endIdx, startIdx + 1))

    for (let i = from; i < to; i++) {
      if (OVERRIDABLE_CODES.has(view[i] ?? '0')) {
        days[i] = override
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
 * Calls `/me/calendar/getSchedule` for a single chunk of mailboxes.
 *
 * Throttling (HTTP 429) is handled by the Graph SDK's default `RetryHandler`
 * middleware, which respects `Retry-After` and backs off exponentially, so
 * no bespoke retry logic is needed here.
 */
async function postGetSchedule(
  client: Client,
  schedules: string[],
  timeWindow: ReturnType<typeof buildTimeWindow>,
  signal?: AbortSignal,
): Promise<GraphScheduleInformation[]> {
  const response = await client
    .api('/me/calendar/getSchedule')
    .header('Prefer', 'outlook.timezone="UTC"')
    .options({ signal })
    .post({
      schedules,
      startTime: timeWindow.startTime,
      endTime: timeWindow.endTime,
      availabilityViewInterval: 1440,
    })

  return (response?.value ?? []) as GraphScheduleInformation[]
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
  client: Client,
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
      results.set(person.id, unavailableSchedule('', dayCount))
      continue
    }

    toFetch.push(person)
  }

  if (toFetch.length === 0) return results

  const timeWindow = buildTimeWindow(year, month)
  const chunks = chunk(toFetch, SCHEDULES_PER_REQUEST)

  await runWithConcurrency(chunks, concurrency, async (peopleChunk) => {
    const mails = peopleChunk.map((person) => person.mail as string)
    let entries: GraphScheduleInformation[]
    try {
      entries = await postGetSchedule(client, mails, timeWindow, options.signal)
    } catch {
      entries = mails.map((mail) => ({
        scheduleId: mail,
        error: { message: 'request failed' },
      }))
    }

    for (let i = 0; i < peopleChunk.length; i++) {
      const person = peopleChunk[i]
      const entry = entries[i] ?? {
        scheduleId: person.mail as string,
        error: { message: 'missing response entry' },
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

  return results
}
