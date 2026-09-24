# Microsoft To Do API Client

Lightweight TypeScript client for Microsoft To Do through Microsoft Graph v1.0.
It uses native `fetch` and local, delegated Microsoft sign-in for the included
examples.

## Local setup

1. In **Microsoft Entra ID → App registrations**, create a [new application registration](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app).
2. Select the account audience you need. For personal plus work/school accounts,
   use the `common` authority.
3. Under **Authentication**, enable **Allow public client flows** for this local
   desktop/CLI app. No redirect URI is needed for device-code flow; see
   [desktop app configuration](https://learn.microsoft.com/en-us/entra/identity-platform/scenario-desktop-app-configuration).
4. Under **API permissions**, add delegated Microsoft Graph `Tasks.ReadWrite`.
5. Copy `.env.example` to `.env`, set `GRAPH_CLIENT_ID` to the registration's
   application (client) ID, and leave `GRAPH_TENANT_ID=common` for the general
   personal-plus-work/school audience.
6. Run `npm run example:auth`. On first use, follow Microsoft's printed
   device-code instructions. Later runs normally use the protected local cache
   silently, until Microsoft requires interaction.

The scripts use Node 24's `--env-file-if-exists=.env` option, so no dotenv
dependency is needed.

## Examples

Authenticate without modifying To Do data:

```sh
npm run example:auth
```

Create or reuse a generic example list, then create one task using the shared
authentication provider:

```sh
npm run example:basic
```

The shopping example is optional convenience functionality and creates its
configured shopping list and task batch:

```sh
npm run example:shopping
```

For troubleshooting only, `GRAPH_ACCESS_TOKEN` may be set in `.env` as an
explicit override. Do not commit it.

If multiple Microsoft accounts are cached, authentication stops. Clear this
app's protected token cache, then sign in with the intended account. On macOS
and Linux, this includes the Keychain/Secret Service entry for service
`microsoft-todo-api`, account `default`; changing the cache file path is not enough.

## Security and cost model

- This app runs locally: it creates no server, database, Azure compute, or
  hosted billing component.
- Entra app registration is free. A verification card can be requested, but it
  is not charged for Entra ID Free.
- Microsoft To Do is free for personal Microsoft accounts. Work/school use is
  governed by that account's Microsoft 365 or Exchange Online license.
- Authentication is delegated to the signed-in user. This is not an unattended
  app-only service and it uses no client secret.
- The MSAL token cache lives under the user's local cache directory (on Windows,
  `%LOCALAPPDATA%\microsoft-todo-api\msal-cache.json`) with current-user
  protection. Windows uses DPAPI and macOS uses Keychain through Microsoft's
  cache extension. Linux requires a Secret Service/LibSecret provider; if it is
  unavailable, authentication stops rather than writing a plaintext cache.
- Tokens and cache contents are never printed or committed.

## References

- [Acquire tokens with MSAL Node](https://learn.microsoft.com/en-us/entra/msal/javascript/node/acquire-token-requests)
- [Microsoft Graph delegated authorization](https://learn.microsoft.com/en-us/graph/auth-v2-user)
- [Create a To Do task](https://learn.microsoft.com/en-us/graph/api/todotasklist-post-tasks?view=graph-rest-1.0)
- [Microsoft Authentication Extensions for Node](https://learn.microsoft.com/en-us/entra/msal/javascript/node/extensions)
- [Microsoft To Do account requirements](https://support.microsoft.com/en-US/Outlook/which-accounts-can-i-use-microsoft-to-do-with)

## Test

```sh
npm test
```
