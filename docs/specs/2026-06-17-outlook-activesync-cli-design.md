# Outlook ActiveSync CLI — Design

**Date:** 2026-06-17
**Status:** Approved
**Binary name:** `outlook`

## 1. Purpose

A command-line tool for reading mail from an Outlook/Exchange server over the
ActiveSync protocol. It lets a user:

1. Create and verify a local configuration (connection settings + credentials).
2. List emails from the last *X* hours (titles, timestamp, id).
3. List emails from the last *X* hours including full bodies.
4. Fetch a single full email by id.

Output is available as human-readable plain text (default) or JSON.

The CLI reuses two existing codebases:

- **`cli-example/`** — the reference for CLI ergonomics: how configuration and
  the user password are handled, and how the `commander`-based CLI is built.
- **`outlook-example/`** — the reference (and source) for the ActiveSync client;
  its `activesync/` implementation is copied largely as-is and adapted.

## 2. Key facts that shaped the design

From reading `outlook-example/activesync/`:

- `ActiveSyncClient.getMailboxMessages()` performs endpoint discovery →
  Provision → FolderSync → paginated Sync, returning **all recent messages
  (up to ~500: `maxPages` 10 × `windowSize` 50) with full bodies already
  included**. There is no server-side time filtering and no fetch-by-id at the
  protocol layer; the `days` argument is ignored inside the client and filtering
  happens locally.
- The Sync request uses `BodyPreference` Type 1 (plain text), so `bodyText` is
  populated for normal messages; HTML-only messages fall back to `preview`.
- `NormalizedMessage` already carries everything the CLI needs:
  `id, subject, from, to[], cc[], receivedAt, preview, bodyText, bodyHtml`.

Consequences:

- Time-window filtering and the merged Inbox+Sent view are done **locally**
  after a sync.
- Reliable fetch-by-id requires a **new** `ItemOperations` Fetch command added to
  the client (the one genuinely new piece of protocol code).

## 3. Decisions

| Topic | Decision |
| --- | --- |
| Listing command shape | One `emails list` command; `--body` flag adds full bodies. Separate `emails get <id>`. |
| Mailboxes | Read **Inbox + Sent**. `list` shows both merged by default; `--mailbox inbox\|sent` narrows. |
| Fetch by id | Implement a real `ItemOperations` Fetch in the client (reliable for any id). |
| Binary name | `outlook` |
| Default time window | 24 hours |
| List row format | `mailbox · receivedAt · id · subject · from` |
| Output | `--output text\|json`, **text default** |

## 4. Project layout

A single TypeScript package mirroring `cli-example`'s structure.

```
src/
  index.ts          # shebang entry → createCli().run()
  cli.ts            # commander program + command wiring
  config.ts         # file config (~/.config/outlook/config.json) + password in keyring
  io.ts             # printOutput (text|json), prompt, formatError  (from cli-example)
  client.ts         # OutlookClient wrapper over ActiveSyncClient (fetch / filter / find)
  format.ts         # plain-text formatters for email list & single email
  types.ts          # mail types (from outlook-example)
  activesync/       # copied from outlook-example, adapted:
    client.ts       #   + ItemOperations Fetch (fetchMessage)
    discovery.ts
    parser.ts       #   + ItemOperations response parsing
    provision.ts
    wbxml.ts        #   verify/extend ItemOperations codepage tokens
    xml.ts
```

Module/runtime conventions (TS, ESM, test runner, lint) follow `cli-example`.

## 5. Configuration

Same split as `cli-example`:

- Non-secret fields in `~/.config/outlook/config.json`.
- Password stored in the OS keyring via `keytar` (service `outlook`, account `cli`).

Config fields:

- `email` — mailbox address (also used to derive the discovery domain).
- `username` — auth username (often same as email; kept separate per the client).
- `host` *(optional)* — drives ActiveSync endpoint discovery.
- `endpoint` *(optional)* — explicit endpoint, bypasses discovery.
- `protocolVersion` *(optional)* — defaults to `14.1`.
- Device identifiers (`deviceId`, `deviceType`, user agent, etc.) keep the
  hard-coded iPhone defaults from `getOutlookEmails`; not user-configurable.

### `config` commands

- **`config init`** — prompts for email, username, password, host; **verifies**
  by running discovery → Provision → FolderSync; on success writes config file +
  saves password to keyring. Mirrors `cli-example`'s `config init` flow
  (prompt → verify via a live call → persist).
- **`config verify`** — re-runs discovery → Provision → FolderSync against the
  saved config and reports success/failure.
- **`config show`** — prints saved config with the password redacted (`****`).

## 6. Email commands

### `emails list [options]`

Syncs Inbox + Sent, filters to the last N hours, prints **titles only**.

Options:

- `--hours <N>` — lookback window in hours (default `24`).
- `--mailbox <inbox|sent>` — restrict to one mailbox (default: both, merged).
- `--body` — include the full plain-text body for each message.
- `--output <text|json>` — default `text`.

Behaviour:

- Messages from both mailboxes are merged and sorted **newest-first**.
- Each text row: `mailbox · receivedAt · id · subject · from`.
- With `--body`, each entry is followed by its body (`bodyText`, falling back to
  `preview` when empty).
- JSON output emits the `NormalizedMailboxMessage` / `NormalizedMessage` shape.

### `emails get <id> [options]`

Fetches a single full email by id.

- Runs FolderSync to resolve the Inbox and Sent collection ids.
- Issues an **`ItemOperations` Fetch** by `ServerId`, trying Inbox then Sent
  until one resolves the item.
- `--output <text|json>` — default `text`. Text shows headers
  (from/to/cc/received/subject) followed by the full body.
- If the id is not found in either collection, exits with a clear error.

## 7. Output & errors

- `printOutput` from `cli-example` handles `text` vs `json`.
  - **text**: human-readable lines (list rows / single-email layout).
  - **json**: `JSON.stringify` of the normalized data.
- Errors are formatted via `formatError` (with `Caused by:` chaining) to stderr;
  the process exits with code `1`. Success exits `0`.
- ActiveSync status / HTTP failures surface the underlying classification
  (`bad-credentials`, `device-blocked`, `endpoint-not-found`, …) already produced
  by `discovery.ts` / the client.

## 8. The new protocol code: `ItemOperations` Fetch

The only substantive new logic beyond copying:

- **Request builder** — `buildItemOperationsFetchXml({ collectionId, serverId, protocolVersion })`
  producing an `ItemOperations:` Fetch with `Store=Mailbox`, `airsync:CollectionId`,
  `airsync:ServerId`, and an `airsyncbase:BodyPreference` Type 1.
- **Client method** — `ActiveSyncClient.fetchMessage({ collectionId, serverId })`
  encodes via `encodeWbxml`, POSTs through `executeCommand("ItemOperations", …)`,
  decodes, and parses out the single message's `ApplicationData`.
- **Parser** — `parseItemOperationsXml(xml)` returning the message in the same
  `NormalizedMessage` shape (reusing `normalizeSyncMessages` helpers).
- **WBXML** — `wbxml.ts` must carry the `ItemOperations` codepage tokens.
  Verify during implementation; extend the codepage tables if missing.

The id shown in `list` output is the raw `ServerId`. `get` does not require the
user to specify a mailbox; it tries both resolved collections.

## 9. Testing

Port `cli-example`'s test setup (vitest). Cover:

- **Hours filtering** — `receivedAt` boundary cases around the cutoff.
- **Text formatters** — list rows and single-email layout (`format.ts`).
- **Config** — read/write round-trip with `keytar` mocked; redaction.
- **ItemOperations** — request builder output and response parsing with a mocked
  `fetchImpl`, following the dependency-injection style the client already uses
  (`fetchImpl`, `createClient`).

## 10. Out of scope (YAGNI)

- Sending, replying, deleting, or moving mail.
- Folders other than Inbox and Sent.
- Attachment download.
- Incremental sync-key persistence between runs (each invocation syncs fresh).
- Multiple stored accounts/profiles.
