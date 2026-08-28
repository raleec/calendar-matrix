import { useEffect, useState } from 'react'
import { getContext } from '@microsoft/power-apps/app'
import type { IContext } from '@microsoft/power-apps/app'
import { isLocalMode, getLocalMe } from '../graph/localClient'

/**
 * Returns the Power Apps host context (user, app, host) or `null` while
 * loading. In local mode (running outside Power Apps) it synthesizes a
 * minimal context from the `/api/me` endpoint so the rest of the app has
 * a consistent identity object regardless of where it is running.
 */
export function usePowerAppsContext(): IContext | null {
  const [context, setContext] = useState<IContext | null>(null)

  useEffect(() => {
    if (isLocalMode()) {
      getLocalMe()
        .then((me) => {
          // Synthesize the minimal IContext shape the app needs.
          setContext({
            user: {
              objectId: me.id,
              fullName: me.displayName,
              userPrincipalName: me.userPrincipalName,
            },
          } as unknown as IContext)
        })
        .catch(() => setContext(null))
      return
    }

    getContext()
      .then(setContext)
      .catch(() => setContext(null))
  }, [])

  return context
}
