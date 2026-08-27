import { useEffect, useState } from 'react'
import { getContext } from '@microsoft/power-apps/app'
import type { IContext } from '@microsoft/power-apps/app'

/**
 * Returns the Power Apps host context (user, app, host) or `null` while
 * loading. Replaces the former MSAL-based `useGraphToken` hook — auth is
 * now handled entirely by the Power Apps host.
 */
export function usePowerAppsContext(): IContext | null {
  const [context, setContext] = useState<IContext | null>(null)

  useEffect(() => {
    getContext()
      .then(setContext)
      .catch(() => setContext(null))
  }, [])

  return context
}
