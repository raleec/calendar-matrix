import { useEffect, useState } from 'react'
import { Office365UsersService } from '../graph/../generated/services/Office365UsersService'
import { usePowerAppsContext } from '../hooks/useGraphToken'

function AccountDisplay({
  name,
  objectId,
}: {
  name: string
  objectId: string
}) {
  const [photoUrl, setPhotoUrl] = useState<string | null>(null)

  useEffect(() => {
    let cancelled = false

    Office365UsersService.UserPhoto_V2(objectId)
      .then((result) => {
        if (!cancelled && result.success && result.data) {
          setPhotoUrl(`data:image/jpeg;base64,${result.data}`)
        }
      })
      .catch(() => {
        // No photo available — fall back to initials.
      })

    return () => {
      cancelled = true
      if (photoUrl) URL.revokeObjectURL(photoUrl)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [objectId])

  return (
    <div className="account">
      {photoUrl ? (
        <img className="account-photo" src={photoUrl} alt="" />
      ) : (
        <span
          className="account-photo account-photo-placeholder"
          aria-hidden="true"
        >
          {name.charAt(0) ?? '?'}
        </span>
      )}
      <span className="account-name">{name}</span>
    </div>
  )
}

export function AccountButton() {
  const context = usePowerAppsContext()

  if (!context?.user.fullName || !context.user.objectId) {
    return <div className="account-button" />
  }

  return (
    <div className="account-button">
      <AccountDisplay
        name={context.user.fullName}
        objectId={context.user.objectId}
      />
    </div>
  )
}
