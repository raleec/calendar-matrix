/**
 * Local-mode API client — used when the app is running outside Power Apps
 * (e.g. `npm run local`). All calls go to the local Express proxy on port
 * 3001, which forwards them to Microsoft Graph using `az` CLI credentials.
 */

import type { GraphGroup, GraphPerson } from './types'
import type { GraphScheduleInformation, TimeWindow } from './schedule'

const BASE = '/api'

/**
 * Returns true when the app is running outside the Power Apps host.
 * Detected by hostname (localhost) or the `VITE_LOCAL_MODE` env flag.
 */
export function isLocalMode(): boolean {
  return (
    import.meta.env.VITE_LOCAL_MODE === 'true' ||
    window.location.hostname === 'localhost' ||
    window.location.hostname === '127.0.0.1'
  )
}

async function apiFetch<T>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(`${BASE}${path}`, init)
  if (!res.ok) {
    const text = await res.text().catch(() => res.statusText)
    throw new Error(`Local API ${path} failed (${res.status}): ${text}`)
  }
  return res.json() as Promise<T>
}

// ---------------------------------------------------------------------------
// Identity
// ---------------------------------------------------------------------------

export interface LocalMe {
  id: string
  displayName: string
  userPrincipalName: string
  mail?: string
}

export function getLocalMe(): Promise<LocalMe> {
  return apiFetch<LocalMe>('/me')
}

// ---------------------------------------------------------------------------
// People
// ---------------------------------------------------------------------------

export async function localSearchUsers(query: string, top = 15): Promise<GraphPerson[]> {
  const data = await apiFetch<{ value: Array<{ id?: string; displayName?: string; mail?: string }> }>(
    `/users/search?q=${encodeURIComponent(query)}&top=${top}`,
  )
  return (data.value ?? [])
    .filter((u): u is typeof u & { id: string } => !!u.id)
    .map((u) => ({
      id: u.id,
      displayName: u.displayName ?? u.mail ?? u.id,
      mail: u.mail ?? null,
    }))
}

export async function localSearchGroups(query: string, top = 15): Promise<GraphGroup[]> {
  const data = await apiFetch<{ value: Array<{ id?: string; displayName?: string }> }>(
    `/groups/search?q=${encodeURIComponent(query)}&top=${top}`,
  )
  return (data.value ?? [])
    .filter((g): g is typeof g & { id: string } => !!g.id)
    .map((g) => ({ id: g.id, displayName: g.displayName ?? g.id }))
}

export async function localGetGroupMembers(groupId: string): Promise<GraphPerson[]> {
  const members: GraphPerson[] = []
  let skipToken: string | undefined

  do {
    const qs = skipToken ? `?skipToken=${encodeURIComponent(skipToken)}` : ''
    // eslint-disable-next-line no-await-in-loop
    const data = await apiFetch<{
      value?: Array<{ id?: string; displayName?: string; mail?: string }>
      '@odata.nextLink'?: string
    }>(`/groups/${groupId}/members${qs}`)
    for (const m of data.value ?? []) {
      if (m.id) members.push({ id: m.id, displayName: m.displayName ?? m.mail ?? m.id, mail: m.mail ?? null })
    }
    const next = data['@odata.nextLink']
    skipToken = next ? new URL(next).searchParams.get('$skiptoken') ?? undefined : undefined
  } while (skipToken)

  return members
}

export async function localGetDirectReports(userId: string): Promise<GraphPerson[]> {
  const data = await apiFetch<{ value?: Array<{ id?: string; displayName?: string; mail?: string }> }>(
    `/users/${userId}/directReports`,
  )
  return (data.value ?? [])
    .filter((u): u is typeof u & { id: string } => !!u.id)
    .map((u) => ({ id: u.id, displayName: u.displayName ?? u.mail ?? u.id, mail: u.mail ?? null }))
}

// ---------------------------------------------------------------------------
// Schedule
// ---------------------------------------------------------------------------

export async function localGetSchedule(
  schedules: string[],
  timeWindow: TimeWindow,
): Promise<GraphScheduleInformation[]> {
  const data = await apiFetch<{ value?: GraphScheduleInformation[] }>('/getSchedule', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ schedules, startTime: timeWindow.startTime, endTime: timeWindow.endTime }),
  })
  return data.value ?? []
}
