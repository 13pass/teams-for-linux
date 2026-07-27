# Graph API Module

Opt-in Microsoft Graph client (`graphApi.enabled`, default `false`). It borrows
the Teams client's own Graph token — it asks the page's React handler for one
(`window.teamsForLinuxReactHandler.acquireToken('https://graph.microsoft.com')`)
— and makes authenticated requests against `https://graph.microsoft.com/v1.0`.

No separate login and no app registration: the calls run with whatever scopes
the first-party Teams client token already carries. That is the key constraint —
see **Scopes** below.

## Components

- **[index.js](index.js)** — `GraphApiClient`: token acquisition/cache, generic
  `makeRequest(endpoint, options)`, and typed helpers.
- **[ipcHandlers.js](ipcHandlers.js)** — `graph-api-*` IPC handlers wrapping the
  client. Channels are allow-listed in `app/security/ipcValidator.js` and bridged
  to the renderer in `app/browser/preload.js` under `electronAPI.graphApi`.

## Helpers

Read/write over the user's own data (profile, calendar, mail, people, send chat)
plus the message-read helpers:

| Method / IPC channel | Graph endpoint | Scope needed |
| --- | --- | --- |
| `getChats` / `graph-api-get-chats` | `GET /me/chats` | `Chat.ReadBasic` |
| `getChatMessages` / `graph-api-get-chat-messages` | `GET /chats/{id}/messages` | `Chat.Read` |
| `getJoinedTeams` / `graph-api-get-joined-teams` | `GET /me/joinedTeams` | `Team.ReadBasic.All` |
| `getChannels` / `graph-api-get-channels` | `GET /teams/{id}/channels` | `Channel.ReadBasic.All` |
| `getChannelMessages` / `graph-api-get-channel-messages` | `GET /teams/{id}/channels/{id}/messages` | `ChannelMessage.Read.All` |
| `probeMessageAccess` / `graph-api-probe-message-access` | samples the above | — |

## Probing scope availability

Because the token is borrowed, you can't know up front which read scopes it
carries. `probeMessageAccess()` samples one chat and one channel (no message
content is returned or logged) and reports a per-area outcome plus a scope hint:

```
[GRAPH_API] Message-access probe {
  chats: { ok: true, status: 200, count: 1 },
  chatMessages: { ok: true, status: 200, count: 1 },
  channelMessages: { ok: false, status: 403, error: '...' },
  hints: { chatRead: 'likely present', channelMessageRead: 'FORBIDDEN (ChannelMessage.Read.All scope missing)' }
}
```

Two ways to run it (after signing in):

- Set `graphApi.probeMessageAccess: true` — it runs once ~20s after Teams
  finishes loading and logs the report.
- From the Teams window DevTools console: `await window.electronAPI.graphApi.probeMessageAccess()`.

## Scopes (the important part)

- **Send** works today (`ChatMessage.Send` is in the Teams token).
- **Chat read** (`Chat.Read`) is often present but tenant-dependent — probe to confirm.
- **Channel message read** (`ChannelMessage.Read.All`) is usually **not** in the
  delegated Teams token and will return `403`. There is no way to widen a
  borrowed token; the durable path is a dedicated Entra app registration with
  the scope granted (and, for channel-message export at scale, Microsoft's
  "protected APIs" model, which may require request-access and metering).

## Compliance

Reading/exporting Teams messages is frequently governed by organisation policy,
and `ChannelMessage.Read.All` needs admin consent. Use with your IT team's
approval in enterprise tenants.
