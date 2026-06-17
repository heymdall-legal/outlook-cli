# outlook

A command-line client for reading Outlook/Exchange mailboxes over ActiveSync.

## Install

```bash
npm install -g outlook-activesync-cli
```

This installs the `outlook` binary globally.

### Development install

```bash
yarn install
yarn build
yarn link   # exposes the `outlook` binary, or run via `node dist/index.js`
```

## Configure

```bash
outlook config init     # prompts for email, username, host, password; verifies; saves
outlook config verify   # re-checks the saved connection
outlook config show     # prints saved config (password redacted)
```

Config is stored in `~/.config/outlook/config.json`; the password is kept in the OS keyring.

## Read mail

```bash
outlook emails list                       # titles from the last 24h (Inbox + Sent)
outlook emails list --hours 48            # last 48h
outlook emails list --mailbox inbox       # Inbox only
outlook emails list --body                # include full bodies
outlook emails get <id>                   # full email by id
outlook emails list --output json         # JSON output (default is text)
```

### Options

- `--hours <N>` — lookback window (default 24).
- `--mailbox <inbox|sent>` — restrict to one mailbox (default both).
- `--body` — include full plain-text bodies in `list`.
- `--max-pages <N>` — safety cap on sync pages per mailbox (default 50). A warning is printed to stderr if the cap is hit while more mail within the window may remain.
- `--output <text|json>` — output format (default text).
