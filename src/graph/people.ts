import type { Client } from '@microsoft/microsoft-graph-client'
import type { GraphGroup, GraphPerson } from './types'

const SEARCH_PAGE_SIZE = 15
const MEMBER_PAGE_SIZE = 999

/**
 * Escapes backslashes and double quotes so a raw query string is safe to
 * embed inside a `$search` quoted phrase.
 */
function toSearchTerm(query: string): string {
  return query.trim().replace(/\\/g, '\\\\').replace(/"/g, '\\"')
}

function toGraphPerson(user: {
  id: string
  displayName?: string | null
  mail?: string | null
}): GraphPerson {
  return {
    id: user.id,
    displayName: user.displayName ?? user.mail ?? user.id,
    mail: user.mail ?? null,
  }
}

/**
 * Typeahead search for people by display name or mail address.
 * `$search` requires the `ConsistencyLevel: eventual` header.
 */
export async function searchUsers(
  client: Client,
  query: string,
  signal?: AbortSignal,
): Promise<GraphPerson[]> {
  const term = toSearchTerm(query)
  if (!term) return []

  const response = await client
    .api('/users')
    .header('ConsistencyLevel', 'eventual')
    .search(`"displayName:${term}" OR "mail:${term}"`)
    .select('id,displayName,mail')
    .top(SEARCH_PAGE_SIZE)
    .options({ signal })
    .get()

  const users: Array<{
    id: string
    displayName?: string | null
    mail?: string | null
  }> = response?.value ?? []

  return users.map(toGraphPerson)
}

/** Typeahead search for groups by display name. */
export async function searchGroups(
  client: Client,
  query: string,
  signal?: AbortSignal,
): Promise<GraphGroup[]> {
  const term = toSearchTerm(query)
  if (!term) return []

  const response = await client
    .api('/groups')
    .header('ConsistencyLevel', 'eventual')
    .search(`"displayName:${term}"`)
    .select('id,displayName')
    .top(SEARCH_PAGE_SIZE)
    .options({ signal })
    .get()

  const groups: Array<{ id: string; displayName?: string | null }> =
    response?.value ?? []

  return groups.map((group) => ({
    id: group.id,
    displayName: group.displayName ?? group.id,
  }))
}

/** Expands a group into its member users, following `@odata.nextLink` pages. */
export async function getGroupMembers(
  client: Client,
  groupId: string,
  signal?: AbortSignal,
): Promise<GraphPerson[]> {
  const members: GraphPerson[] = []

  let response = await client
    .api(`/groups/${groupId}/members`)
    .select('id,displayName,mail')
    .top(MEMBER_PAGE_SIZE)
    .options({ signal })
    .get()

  for (;;) {
    const page: Array<{
      id: string
      displayName?: string | null
      mail?: string | null
      '@odata.type'?: string
    }> = response?.value ?? []

    for (const item of page) {
      // Skip nested groups/contacts and only keep user-type members.
      if (
        item['@odata.type'] &&
        item['@odata.type'] !== '#microsoft.graph.user'
      ) {
        continue
      }
      members.push(toGraphPerson(item))
    }

    const nextLink: string | undefined = response?.['@odata.nextLink']
    if (!nextLink) break

    response = await client.api(nextLink).options({ signal }).get()
  }

  return members
}

/** Direct reports of the signed-in user. */
export async function getDirectReports(
  client: Client,
  signal?: AbortSignal,
): Promise<GraphPerson[]> {
  const response = await client
    .api('/me/directReports')
    .select('id,displayName,mail')
    .top(MEMBER_PAGE_SIZE)
    .options({ signal })
    .get()

  const reports: Array<{
    id: string
    displayName?: string | null
    mail?: string | null
  }> = response?.value ?? []

  return reports.map(toGraphPerson)
}
