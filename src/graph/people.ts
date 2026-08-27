/**
 * Graph-backed helpers for searching people and groups, then normalizing the
 * connector responses into the app's shared person/group shapes.
 */
import { Office365GroupsService } from '../generated/services/Office365GroupsService'
import { Office365UsersService } from '../generated/services/Office365UsersService'
import type { GraphGroup, GraphPerson } from './types'

const SEARCH_PAGE_SIZE = 15
const MEMBER_PAGE_SIZE = 999

/**
 * Normalizes connector user payloads into a single {@link GraphPerson} shape.
 *
 * The Office365Users connector returns PascalCase fields, while group-member
 * responses from the Groups connector use camelCase, so this helper accepts
 * either form.
 */
function userToPerson(user: {
  Id?: string; id?: string
  DisplayName?: string; displayName?: string
  Mail?: string; mail?: string
}): GraphPerson {
  const id = (user.Id ?? user.id) || ''
  return {
    id,
    displayName: user.DisplayName ?? user.displayName ?? user.Mail ?? user.mail ?? id,
    mail: user.Mail ?? user.mail ?? null,
  }
}

/** Typeahead search for people by display name or email. */
export async function searchUsers(query: string): Promise<GraphPerson[]> {
  const term = query.trim()
  if (!term) return []
  const result = await Office365UsersService.SearchUser(term, SEARCH_PAGE_SIZE)
  if (!result.success || !result.data) return []
  return result.data.map(userToPerson)
}

/** Typeahead search for groups by display name (prefix match). */
export async function searchGroups(query: string): Promise<GraphGroup[]> {
  const term = query.trim()
  if (!term) return []
  const safe = term.replace(/'/g, "''")
  const result = await Office365GroupsService.ListGroups(
    undefined,
    undefined,
    `startswith(displayName,'${safe}')`,
    SEARCH_PAGE_SIZE,
  )
  if (!result.success || !result.data?.value) return []
  return result.data.value
    .filter((g): g is NonNullable<typeof g> & { id: string } => !!g.id)
    .map((g) => ({ id: g.id, displayName: g.displayName ?? g.id }))
}

/** Expands a group into its member users, paging with skipToken. */
export async function getGroupMembers(groupId: string): Promise<GraphPerson[]> {
  const members: GraphPerson[] = []
  let skipToken: string | undefined
  do {
    // eslint-disable-next-line no-await-in-loop
    const result = await Office365GroupsService.ListGroupMembers(groupId, MEMBER_PAGE_SIZE)
    if (!result.success) break
    for (const m of result.data?.value ?? []) {
      if (m.id) members.push(userToPerson(m))
    }
    skipToken = result.skipToken
  } while (skipToken)
  return members
}

/** Direct reports of the signed-in user, identified by their Entra object ID. */
export async function getDirectReports(userObjectId: string): Promise<GraphPerson[]> {
  const result = await Office365UsersService.DirectReports_V2(userObjectId)
  if (!result.success || !result.data?.value) return []
  return result.data.value
    .filter((u): u is NonNullable<typeof u> & { id: string } => !!u.id)
    .map((u) => ({ id: u.id, displayName: u.displayName ?? u.mail ?? u.id, mail: u.mail ?? null }))
}
