import { useCallback } from 'react'
import { useMsal } from '@azure/msal-react'
import {
  InteractionRequiredAuthError,
  type AccountInfo,
} from '@azure/msal-browser'
import { loginRequest } from '../authConfig'

/**
 * Returns a function that acquires a Microsoft Graph access token for the
 * signed-in account, trying a silent (cached/refresh) acquisition first and
 * only falling back to an interactive popup when silent acquisition fails
 * with `InteractionRequiredAuthError`.
 */
export function useGraphToken() {
  const { instance, accounts } = useMsal()

  return useCallback(
    async (account?: AccountInfo) => {
      const activeAccount = account ?? accounts[0]
      if (!activeAccount) {
        throw new Error(
          'No signed-in account available to acquire a token for.',
        )
      }

      const request = { ...loginRequest, account: activeAccount }

      try {
        const result = await instance.acquireTokenSilent(request)
        return result.accessToken
      } catch (error) {
        if (error instanceof InteractionRequiredAuthError) {
          const result = await instance.acquireTokenPopup(request)
          return result.accessToken
        }
        throw error
      }
    },
    [instance, accounts],
  )
}
