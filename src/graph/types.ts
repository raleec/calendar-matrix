/**
 * A person that can be added to the calendar matrix, whether picked
 * directly or expanded from a group's membership.
 */
export interface GraphPerson {
  id: string
  displayName: string
  mail: string | null
}

/**
 * A Microsoft Entra ID (Azure AD) group returned from a directory search.
 */
export interface GraphGroup {
  id: string
  displayName: string
}
