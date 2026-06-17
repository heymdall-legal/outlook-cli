---
name: outlook-cli
description: Read Outlook/Exchange mailboxes from the terminal with the `outlook` CLI (outlook-activesync-cli). Use when a user wants to list or fetch their Outlook/Exchange emails, check recent mail, search their inbox or sent folder, read an email by id, or set up and verify the Outlook connection over ActiveSync.
---

# Outlook CLI

Use this skill to read Outlook/Exchange mail over ActiveSync with the `outlook` binary (package `outlook-activesync-cli`). Mail is **read-only** — there is no command to send, reply, delete, or move messages.

## Prerequisites

1. Check the CLI is installed: run `outlook --help`.
2. If `outlook: command not found`, install it: `npm install -g outlook-activesync-cli`.
3. Check the connection is configured: run `outlook config show`. If it prints config, you are ready. If it errors with "Missing required configuration", the user must run `outlook config init` (interactive — it prompts for email, username, host, password). **Never run `config init` yourself and never ask the user for their password**; tell them to run it.

Config lives in `~/.config/outlook/config.json`; the password is stored in the OS keyring (keytar), not in the file.

## Reading mail

Always prefer `--output json` when you need to parse results or feed them into other steps; use the default text output when showing results directly to the user.

```bash
outlook emails list                    # subjects from the last 24h (Inbox + Sent)
outlook emails list --hours 48         # widen the lookback window
outlook emails list --mailbox inbox    # Inbox only (or --mailbox sent)
outlook emails list --body             # include full plain-text bodies
outlook emails list --output json      # machine-readable output
outlook emails get <id>                # one full email by id
outlook emails get <id> --output json  # full email as JSON
```

### Typical flow

1. Run `outlook emails list` (optionally `--output json`) to get rows. Each row carries an `id`.
2. To read a specific message in full, take its `id` and run `outlook emails get <id>`.

`get` searches both Inbox and Sent; an unknown id is an error.

## Options reference

- `--hours <N>` — lookback window in hours (default `24`). Must be a positive integer.
- `--mailbox <inbox|sent>` — restrict to one mailbox; default is both.
- `--body` — (`list` only) include full plain-text bodies, not just subject lines.
- `--max-pages <N>` — safety cap on sync pages per mailbox (default `50`). If the cap is hit while more mail in the window may remain, a warning is printed to **stderr**; raise this if you suspect results were truncated.
- `--output <text|json>` — output format (default `text`), available on both `list` and `get`.

## Config commands

- `outlook config init` — interactive setup; prompts, verifies the connection, then saves. User-run only.
- `outlook config verify` — re-check the saved connection against the server.
- `outlook config show` — print saved config as JSON with the password redacted (`****`).

## Output shapes

Text `list` rows are ` · `-separated: `mailbox · receivedAt · id · subject · from`.

JSON output is an array (for `list`) or a single object (for `get`) of:

```json
{
  "mailbox": "inbox",
  "message": {
    "id": "...",
    "subject": "...",
    "from": "...",
    "to": ["..."],
    "cc": ["..."],
    "receivedAt": "...",
    "preview": "...",
    "bodyText": "...",
    "bodyHtml": "..."
  }
}
```

## Notes

- Exit code is `0` on success, `1` on failure; errors go to stderr and may include a `Caused by:` chain.
- "No emails found in the requested window." means the query worked but matched nothing — try a larger `--hours`.
- If `keytar`/keyring errors appear, the user may need libsecret installed (Linux) and to reinstall.
