import { Client } from '@microsoft/microsoft-graph-client'

/**
 * Creates a Microsoft Graph `Client` whose auth provider delegates token
 * acquisition to the supplied callback (backed by MSAL — see
 * `useGraphToken`). The client is intentionally re-created whenever the
 * token callback changes so consumers don't need to manage the Graph SDK
 * lifecycle themselves.
 */
export function createGraphClient(getAccessToken: () => Promise<string>) {
  return Client.init({
    authProvider: async (done) => {
      try {
        const token = await getAccessToken()
        done(null, token)
      } catch (error) {
        done(error, null)
      }
    },
  })
}
