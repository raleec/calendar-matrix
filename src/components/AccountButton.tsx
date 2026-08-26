import { useEffect, useState } from 'react'
import {
  AuthenticatedTemplate,
  UnauthenticatedTemplate,
  useMsal,
} from '@azure/msal-react'
import type { AccountInfo } from '@azure/msal-browser'
import { loginRequest, graphConfig } from '../authConfig'
import { useGraphToken } from '../hooks/useGraphToken'

function SignedIn({ account }: { account: AccountInfo }) {
  const { instance } = useMsal()
  const getGraphToken = useGraphToken()
  const [photoUrl, setPhotoUrl] = useState<string | null>(null)

  useEffect(() => {
    let objectUrl: string | null = null
    let cancelled = false

    async function loadPhoto() {
      try {
        const accessToken = await getGraphToken(account)
        const response = await fetch(graphConfig.photoEndpoint, {
          headers: { Authorization: 'Bearer ' + accessToken },
        })
        if (!response.ok || cancelled) return
        const blob = await response.blob()
        objectUrl = URL.createObjectURL(blob)
        if (!cancelled) setPhotoUrl(objectUrl)
      } catch {
        // No photo available (or permission denied) — fall back to initials.
      }
    }

    void loadPhoto()

    return () => {
      cancelled = true
      if (objectUrl) URL.revokeObjectURL(objectUrl)
    }
  }, [account, getGraphToken])

  const signOut = () => {
    void instance.logoutRedirect({ account })
  }

  return (
    <div className="account">
      {photoUrl ? (
        <img className="account-photo" src={photoUrl} alt="" />
      ) : (
        <span
          className="account-photo account-photo-placeholder"
          aria-hidden="true"
        >
          {account.name?.charAt(0) ?? '?'}
        </span>
      )}
      <span className="account-name">{account.name}</span>
      <button type="button" className="account-signout" onClick={signOut}>
        Sign out
      </button>
    </div>
  )
}

export function AccountButton() {
  const { instance, accounts } = useMsal()

  const signIn = () => {
    void instance.loginRedirect(loginRequest)
  }

  return (
    <div className="account-button">
      <AuthenticatedTemplate>
        {accounts[0] && <SignedIn account={accounts[0]} />}
      </AuthenticatedTemplate>
      <UnauthenticatedTemplate>
        <button type="button" className="account-signin" onClick={signIn}>
          Sign in
        </button>
      </UnauthenticatedTemplate>
    </div>
  )
}
