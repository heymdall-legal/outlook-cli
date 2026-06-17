**Goal:** Build `outlook`, a TypeScript CLI that reads mail from an Exchange/Outlook server over ActiveSync — configure & verify a connection, list emails from the last N hours (titles or full bodies), and fetch a single email by id, with text (default) or JSON output.

**Architecture:** Copy the proven ActiveSync protocol layer from `outlook-example/activesync/` into `src/activesync/` and extend it with an `ItemOperations` Fetch command (the only new protocol code). A thin `OutlookClient` wrapper handles time-window filtering and mailbox selection. The CLI layer (`commander`) and config/IO helpers follow `cli-example`'s patterns: non-secret config in `~/.config/outlook/config.json`, password in the OS keyring via `keytar`.

**Tech Stack:** Node.js (ESM, NodeNext), TypeScript, `commander`, `keytar`, `vitest`, `tsx`, Yarn 4.

---

## Conventions for the implementing engineer

- **Run tests:** `yarn test <path>` runs vitest once on a file (e.g. `yarn test src/activesync/parser.test.ts`). `yarn test` runs the whole suite.
- **Build:** `yarn build` (runs `tsc`). **Type-check only:** `yarn typecheck`.
- **ESM import extensions are mandatory.** Because `tsconfig` uses `moduleResolution: "NodeNext"`, every relative import MUST end in `.js` even though the source file is `.ts` (e.g. `import { x } from "./parser.js"`). This matches the existing `outlook-example` files.
- **Commit after every task** with the exact message given. Never commit `node_modules/` or `dist/`.
- The two example directories (`cli-example/`, `outlook-example/`) are **reference only** — never import from them; copy what you need into `src/`.

## File structure (built across the tasks)

```
package.json            # deps, scripts, bin: { outlook }   (Task 1)
tsconfig.json           # NodeNext ESM, strict             (Task 1)
vitest.config.ts        # node env                          (Task 1)
.gitignore              # node_modules, dist                (Task 1)
src/
  index.ts              # shebang entry → createCli().run() (Task 11)
  cli.ts                # commander program + wiring        (Task 11)
  config.ts             # file config + keyring password    (Task 7)
  io.ts                 # printOutput / prompt / formatError (Task 8)
  format.ts             # plain-text email formatters       (Task 9)
  client.ts             # OutlookClient wrapper             (Task 10)
  types.ts              # mail types (copied + sent rename) (Task 2)
  messages.ts           # time-window filters               (Task 2)
  activesync/
    client.ts           # ActiveSyncClient (+ fetch-by-id)  (Task 2, 5, 6)
    discovery.ts        # endpoint discovery (copied)        (Task 2)
    parser.ts           # XML parsers (+ ItemOperations)     (Task 2, 4)
    provision.ts        # provisioning (copied)              (Task 2)
    wbxml.ts            # WBXML codec (+ ItemOperations cp)  (Task 2, 3)
    xml.ts              # regex XML helpers (copied)         (Task 2)
```

---

## Task 1: Project scaffolding & toolchain

**Files:**
- Create: `package.json` (overwrite the 2-line stub), `tsconfig.json`, `vitest.config.ts`, `.gitignore`, `src/smoke.test.ts`

- [ ] **Step 1: Write `package.json`**

```json
{
  "name": "outlook-activesync-cli",
  "version": "0.1.0",
  "packageManager": "yarn@4.17.0",
  "type": "module",
  "bin": {
    "outlook": "./dist/index.js"
  },
  "scripts": {
    "build": "tsc",
    "typecheck": "tsc --noEmit",
    "test": "vitest run",
    "dev": "tsx src/index.ts"
  },
  "dependencies": {
    "commander": "^12.1.0",
    "keytar": "^7.9.0"
  },
  "devDependencies": {
    "@types/node": "^22.7.0",
    "tsx": "^4.19.0",
    "typescript": "^5.6.0",
    "vitest": "^2.1.0"
  }
}
```

- [ ] **Step 2: Write `tsconfig.json`**

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "outDir": "dist",
    "rootDir": "src",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "resolveJsonModule": true,
    "declaration": false,
    "sourceMap": false
  },
  "include": ["src"],
  "exclude": ["node_modules", "dist", "cli-example", "outlook-example"]
}
```

- [ ] **Step 3: Write `vitest.config.ts`**

```ts
import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    include: ["src/**/*.test.ts"],
  },
});
```

- [ ] **Step 4: Write `.gitignore`**

```
node_modules/
dist/
*.log
```

- [ ] **Step 5: Write a smoke test at `src/smoke.test.ts`**

```ts
import { describe, expect, it } from "vitest";

describe("toolchain", () => {
  it("runs vitest", () => {
    expect(1 + 1).toBe(2);
  });
});
```

- [ ] **Step 6: Install dependencies**

Run: `yarn install`
Expected: completes; creates `node_modules/`. (If `keytar` native build fails on this machine, note it — it is only needed at runtime; tests mock it. Proceed.)

- [ ] **Step 7: Run the smoke test**

Run: `yarn test src/smoke.test.ts`
Expected: PASS (1 test).

- [ ] **Step 8: Verify type-check passes**

Run: `yarn typecheck`
Expected: no errors.

- [ ] **Step 9: Commit**

```bash
git add package.json tsconfig.json vitest.config.ts .gitignore src/smoke.test.ts yarn.lock
git commit -m "chore: scaffold outlook CLI toolchain"
```

---

## Task 2: Copy the ActiveSync protocol layer (with outbox→sent rename)

Copy six protocol files + `types.ts` + `messages.ts` from `outlook-example/` into `src/`, adjusting one type and one string literal so the "Sent Items" mailbox is named `sent` everywhere (spec §5.1). Add an hours-based filter to `messages.ts`.

**Files:**
- Create: `src/activesync/wbxml.ts`, `src/activesync/xml.ts`, `src/activesync/discovery.ts`, `src/activesync/provision.ts`, `src/activesync/parser.ts`, `src/activesync/client.ts`, `src/types.ts`, `src/messages.ts`
- Test: `src/activesync/parser.test.ts`, `src/messages.test.ts`

- [ ] **Step 1: Copy the files verbatim**

```bash
cp outlook-example/activesync/wbxml.ts src/activesync/wbxml.ts
cp outlook-example/activesync/xml.ts src/activesync/xml.ts
cp outlook-example/activesync/discovery.ts src/activesync/discovery.ts
cp outlook-example/activesync/provision.ts src/activesync/provision.ts
cp outlook-example/activesync/parser.ts src/activesync/parser.ts
cp outlook-example/activesync/client.ts src/activesync/client.ts
cp outlook-example/types.ts src/types.ts
cp outlook-example/messages.ts src/messages.ts
```

- [ ] **Step 2: Rename the mailbox kind in `src/types.ts`**

Change line 13 from:

```ts
export type MailboxKind = "inbox" | "outbox";
```

to:

```ts
export type MailboxKind = "inbox" | "sent";
```

- [ ] **Step 3: Update the literal in `src/activesync/client.ts`**

In `getMailboxMessages`, the `outboxMessages` block calls `syncFolderMessages` with `mailbox: "outbox"`. Change that one literal to `mailbox: "sent"`:

```ts
    const outboxMessages = outbox
      ? await this.syncFolderMessages({
          mailbox: "sent",
          collectionId: outbox.serverId,
          maxPages,
          windowSize,
        })
      : [];
```

(Leave the local variable name `outbox` and the function `findOutboxFolder` as-is — they still resolve the Sent Items folder; only the user-facing `MailboxKind` value changes.)

- [ ] **Step 4: Add an hours-based filter to `src/messages.ts`**

Append this function (keep the existing `filterRecentMessages` / `filterMessagesSince`):

```ts
export function filterMessagesWithinHours(
  messages: NormalizedMailboxMessage[],
  hours: number,
  now: Date = new Date(),
): NormalizedMailboxMessage[] {
  const cutoff = new Date(now.getTime() - hours * 60 * 60 * 1000);

  return messages.filter((entry) => {
    const receivedAt = new Date(entry.message.receivedAt);
    return !Number.isNaN(receivedAt.valueOf()) && receivedAt >= cutoff;
  });
}
```

- [ ] **Step 5: Write the parser/round-trip test at `src/activesync/parser.test.ts`**

```ts
import { describe, expect, it } from "vitest";

import { decodeWbxml, encodeWbxml } from "./wbxml.js";
import {
  buildFolderSyncRequestXml,
  findInboxFolder,
  findOutboxFolder,
  normalizeSyncMessages,
  parseFolderSyncXml,
  parseSyncXml,
} from "./parser.js";

describe("FolderSync parsing", () => {
  const sample =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<FolderSync xmlns="FolderHierarchy:"><Status>1</Status><SyncKey>1</SyncKey>' +
    "<Changes>" +
    "<Add><ServerId>2</ServerId><ParentId>0</ParentId><DisplayName>Inbox</DisplayName><Type>2</Type></Add>" +
    "<Add><ServerId>5</ServerId><ParentId>0</ParentId><DisplayName>Sent Items</DisplayName><Type>5</Type></Add>" +
    "</Changes></FolderSync>";

  it("extracts folders and finds inbox + sent", () => {
    const parsed = parseFolderSyncXml(sample);
    expect(parsed.status).toBe("1");
    expect(parsed.folders).toHaveLength(2);
    expect(findInboxFolder(parsed.folders)?.serverId).toBe("2");
    expect(findOutboxFolder(parsed.folders)?.serverId).toBe("5");
  });
});

describe("Sync message normalization", () => {
  it("maps a plain-text email into NormalizedMessage", () => {
    const sample =
      '<?xml version="1.0" encoding="utf-8"?>' +
      '<Sync xmlns="AirSync:"><Collections><Collection><SyncKey>2</SyncKey><Status>1</Status>' +
      "<Commands><Add><ServerId>2:1</ServerId><ApplicationData>" +
      "<Subject>Hi</Subject><From>a@b.com</From><DateReceived>2026-06-17T09:00:00.000Z</DateReceived>" +
      "<Type>1</Type><Data>Body</Data>" +
      "</ApplicationData></Add></Commands></Collection></Collections></Sync>";
    const parsed = parseSyncXml(sample);
    const [message] = normalizeSyncMessages(parsed.messages);
    expect(message.id).toBe("2:1");
    expect(message.subject).toBe("Hi");
    expect(message.bodyText).toBe("Body");
  });
});

describe("WBXML round-trip", () => {
  it("encodes and decodes a FolderSync request without losing structure", () => {
    const xml = buildFolderSyncRequestXml("0");
    const decoded = decodeWbxml(encodeWbxml(xml));
    expect(decoded).toContain("<FolderSync");
    expect(decoded).toContain("<SyncKey>0</SyncKey>");
  });
});
```

- [ ] **Step 6: Write the messages test at `src/messages.test.ts`**

```ts
import { describe, expect, it } from "vitest";

import { filterMessagesWithinHours } from "./messages.js";
import type { NormalizedMailboxMessage } from "./types.js";

function entry(receivedAt: string): NormalizedMailboxMessage {
  return {
    mailbox: "inbox",
    message: {
      id: receivedAt,
      subject: "s",
      from: "f",
      to: [],
      cc: [],
      receivedAt,
      preview: "",
      bodyText: "",
      bodyHtml: "",
    },
  };
}

describe("filterMessagesWithinHours", () => {
  const now = new Date("2026-06-17T12:00:00.000Z");

  it("keeps messages inside the window and drops older ones", () => {
    const messages = [
      entry("2026-06-17T11:00:00.000Z"), // 1h ago -> keep
      entry("2026-06-16T11:00:00.000Z"), // 25h ago -> drop
    ];
    const result = filterMessagesWithinHours(messages, 24, now);
    expect(result.map((e) => e.message.id)).toEqual(["2026-06-17T11:00:00.000Z"]);
  });

  it("drops messages with unparseable dates", () => {
    const result = filterMessagesWithinHours([entry("not-a-date")], 24, now);
    expect(result).toHaveLength(0);
  });
});
```

- [ ] **Step 7: Run the tests**

Run: `yarn test src/activesync/parser.test.ts src/messages.test.ts`
Expected: PASS (5 tests).

- [ ] **Step 8: Type-check**

Run: `yarn typecheck`
Expected: no errors.

- [ ] **Step 9: Commit**

```bash
git add src/activesync src/types.ts src/messages.ts src/messages.test.ts src/activesync/parser.test.ts
git commit -m "feat: copy ActiveSync protocol layer with sent-mailbox naming and hours filter"
```

---

## Task 3: Populate the ItemOperations WBXML code page + fix airsync prefix

The copied `wbxml.ts` registers `ItemOperations` as an **empty** code page (`createCodePage("ItemOperations", {})` at ~line 290), so `ItemOperations` requests/responses cannot be encoded or decoded. Populate it with the authoritative tokens, and fix prefix resolution so the `airsync:` prefix maps to the `AirSync` namespace (needed because `ItemOperations` Fetch nests `AirSync` elements).

**Files:**
- Modify: `src/activesync/wbxml.ts`
- Test: `src/activesync/wbxml.test.ts`

- [ ] **Step 1: Populate the ItemOperations code page**

Replace `createCodePage("ItemOperations", {}),` with:

```ts
  createCodePage("ItemOperations", {
    ItemOperations: 0x05,
    Fetch: 0x06,
    Store: 0x07,
    Options: 0x08,
    Range: 0x09,
    Total: 0x0a,
    Properties: 0x0b,
    Data: 0x0c,
    Status: 0x0d,
    Response: 0x0e,
    Version: 0x0f,
    Schema: 0x10,
    Part: 0x11,
    EmptyFolderContents: 0x12,
    DeleteSubFolders: 0x13,
    UserName: 0x14,
    Password: 0x15,
    Move: 0x16,
    DstFldId: 0x17,
    ConversationId: 0x18,
    MoveAlways: 0x19,
  }),
```

(These are code page 20 / `0x14` tokens from `[MS-ASWBXML]`; the array position is already index 20, matching the code page number.)

- [ ] **Step 2: Fix `prefixToNamespace` for the `airsync` prefix**

Find:

```ts
function prefixToNamespace(prefix) {
  if (prefix === "airsyncbase") {
    return "AirSyncBase";
  }

  return prefix.charAt(0).toUpperCase() + prefix.slice(1);
}
```

Replace with:

```ts
function prefixToNamespace(prefix) {
  if (prefix === "airsyncbase") {
    return "AirSyncBase";
  }

  if (prefix === "airsync") {
    return "AirSync";
  }

  return prefix.charAt(0).toUpperCase() + prefix.slice(1);
}
```

(Without this, `airsync:CollectionId` would resolve to the non-existent namespace `Airsync` and throw `Unsupported WBXML namespace`.)

- [ ] **Step 3: Write the test at `src/activesync/wbxml.test.ts`**

```ts
import { describe, expect, it } from "vitest";

import { decodeWbxml, encodeWbxml } from "./wbxml.js";

describe("ItemOperations WBXML", () => {
  it("round-trips an ItemOperations Fetch request with nested AirSync/AirSyncBase elements", () => {
    const xml =
      '<?xml version="1.0" encoding="utf-8"?>' +
      '<ItemOperations xmlns="ItemOperations:" xmlns:airsync="AirSync:" xmlns:airsyncbase="AirSyncBase:">' +
      "<Fetch><Store>Mailbox</Store>" +
      "<airsync:CollectionId>2</airsync:CollectionId>" +
      "<airsync:ServerId>2:42</airsync:ServerId>" +
      "<Options><airsyncbase:BodyPreference><airsyncbase:Type>1</airsyncbase:Type>" +
      "<airsyncbase:TruncationSize>200000</airsyncbase:TruncationSize></airsyncbase:BodyPreference></Options>" +
      "</Fetch></ItemOperations>";

    const decoded = decodeWbxml(encodeWbxml(xml));

    expect(decoded).toContain("<ItemOperations");
    expect(decoded).toContain("<Store>Mailbox</Store>");
    expect(decoded).toContain("<CollectionId>2</CollectionId>");
    expect(decoded).toContain("<ServerId>2:42</ServerId>");
    expect(decoded).toContain("<Type>1</Type>");
  });
});
```

- [ ] **Step 4: Run the test**

Run: `yarn test src/activesync/wbxml.test.ts`
Expected: PASS (1 test).

- [ ] **Step 5: Commit**

```bash
git add src/activesync/wbxml.ts src/activesync/wbxml.test.ts
git commit -m "feat: populate ItemOperations WBXML code page and fix airsync prefix"
```

---

## Task 4: ItemOperations Fetch request builder + response parser

**Files:**
- Modify: `src/activesync/parser.ts`
- Test: `src/activesync/item-operations.test.ts`

- [ ] **Step 1: Add the request builder + response parser to `src/activesync/parser.ts`**

Append at the end of the file:

```ts
export function buildItemOperationsFetchXml({
  protocolVersion = "14.1",
  collectionId,
  serverId,
  truncationSize = 200000,
}: {
  protocolVersion?: string;
  collectionId: string;
  serverId: string;
  truncationSize?: number;
}): string {
  void protocolVersion;

  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<ItemOperations xmlns="ItemOperations:" xmlns:airsync="AirSync:" xmlns:airsyncbase="AirSyncBase:">' +
    "<Fetch><Store>Mailbox</Store>" +
    `<airsync:CollectionId>${escapeXml(collectionId)}</airsync:CollectionId>` +
    `<airsync:ServerId>${escapeXml(serverId)}</airsync:ServerId>` +
    "<Options><airsyncbase:BodyPreference>" +
    "<airsyncbase:Type>1</airsyncbase:Type>" +
    `<airsyncbase:TruncationSize>${truncationSize}</airsyncbase:TruncationSize>` +
    "</airsyncbase:BodyPreference></Options>" +
    "</Fetch></ItemOperations>"
  );
}

export interface ItemOperationsResult {
  status: string;
  message: ParsedSyncMessage | null;
}

export function parseItemOperationsXml(xml: string): ItemOperationsResult {
  const topStatus = getFirstTagText(xml, "Status");
  const fetchBlock = getAllTagBlocks(xml, "Fetch")[0] ?? "";

  if (!fetchBlock) {
    return { status: topStatus, message: null };
  }

  const fetchStatus = getFirstTagText(fetchBlock, "Status") || topStatus;
  const properties = getAllTagBlocks(fetchBlock, "Properties")[0] ?? "";
  const serverId = getFirstTagText(fetchBlock, "ServerId");

  if (!serverId && !properties) {
    return { status: fetchStatus, message: null };
  }

  return {
    status: fetchStatus,
    message: {
      serverId,
      applicationData: {
        subject: getFirstTagText(properties, "Subject"),
        from: getFirstTagText(properties, "From"),
        to: getFirstTagText(properties, "To"),
        cc: getFirstTagText(properties, "Cc"),
        dateReceived: getFirstTagText(properties, "DateReceived"),
        bodyType: getFirstTagText(properties, "Type"),
        bodyData: getFirstTagText(properties, "Data"),
        bodyPreview: getFirstTagText(properties, "Preview"),
      },
    },
  };
}
```

- [ ] **Step 2: Write the test at `src/activesync/item-operations.test.ts`**

```ts
import { describe, expect, it } from "vitest";

import {
  buildItemOperationsFetchXml,
  normalizeSyncMessages,
  parseItemOperationsXml,
} from "./parser.js";

describe("buildItemOperationsFetchXml", () => {
  it("includes the collection id, server id, and a plain-text body preference", () => {
    const xml = buildItemOperationsFetchXml({ collectionId: "2", serverId: "2:42" });
    expect(xml).toContain("<airsync:CollectionId>2</airsync:CollectionId>");
    expect(xml).toContain("<airsync:ServerId>2:42</airsync:ServerId>");
    expect(xml).toContain("<airsyncbase:Type>1</airsyncbase:Type>");
  });
});

describe("parseItemOperationsXml", () => {
  it("parses a successful Fetch response into a normalizable message", () => {
    const xml =
      '<?xml version="1.0" encoding="utf-8"?>' +
      '<ItemOperations xmlns="ItemOperations:"><Status>1</Status><Response><Fetch>' +
      "<Status>1</Status><CollectionId>2</CollectionId><ServerId>2:42</ServerId>" +
      "<Properties><Subject>Hello</Subject><From>a@b.com</From>" +
      "<DateReceived>2026-06-17T09:00:00.000Z</DateReceived>" +
      "<Body><Type>1</Type><Data>Body text</Data></Body></Properties>" +
      "</Fetch></Response></ItemOperations>";

    const result = parseItemOperationsXml(xml);
    expect(result.status).toBe("1");
    expect(result.message?.serverId).toBe("2:42");

    const [normalized] = normalizeSyncMessages([result.message!]);
    expect(normalized.subject).toBe("Hello");
    expect(normalized.bodyText).toBe("Body text");
  });

  it("returns a null message when the item is absent", () => {
    const xml =
      '<?xml version="1.0" encoding="utf-8"?>' +
      '<ItemOperations xmlns="ItemOperations:"><Status>1</Status><Response></Response></ItemOperations>';
    const result = parseItemOperationsXml(xml);
    expect(result.message).toBeNull();
  });
});
```

- [ ] **Step 3: Run the test**

Run: `yarn test src/activesync/item-operations.test.ts`
Expected: PASS (3 tests).

- [ ] **Step 4: Commit**

```bash
git add src/activesync/parser.ts src/activesync/item-operations.test.ts
git commit -m "feat: add ItemOperations Fetch request builder and response parser"
```

---

## Task 5: Fetch-by-id in ActiveSyncClient (`fetchMessage` + `getMessageById`)

Refactor the FolderSync+provision logic out of `getMailboxMessages` into a reusable method, then add fetch-by-id that tries Inbox then Sent.

**Files:**
- Modify: `src/activesync/client.ts`
- Test: `src/activesync/client.test.ts`

- [ ] **Step 1: Add imports for the new parser functions**

In `src/activesync/client.ts`, extend the existing import from `./parser.js` to include the new symbols:

```ts
import {
  buildFolderSyncRequestXml,
  buildItemOperationsFetchXml,
  buildSyncRequestXml,
  findInboxFolder,
  findOutboxFolder,
  normalizeSyncMessages,
  parseFolderSyncXml,
  parseItemOperationsXml,
  parseSyncXml,
  type FolderRecord,
} from "./parser.js";
```

- [ ] **Step 2: Extract `resolveMailFolders` and refactor `getMailboxMessages`**

Replace the body of `getMailboxMessages` (the FolderSync + status-handling block, down to where it computes `inbox`/`outbox`) by delegating to a new method. Add this new method and update `getMailboxMessages` to use it:

```ts
  async resolveMailFolders(reprovisionAttempts = 0): Promise<{
    inbox: FolderRecord;
    sent: FolderRecord | null;
  }> {
    await this.ensureEndpoint();

    const folderSyncXml = await this.executeCommand("FolderSync", buildFolderSyncRequestXml("0"));
    const folderSync = parseFolderSyncXml(folderSyncXml);

    if (folderSync.status === "142" || folderSync.status === "144") {
      if (reprovisionAttempts >= 1) {
        throw new Error(
          `FolderSync failed with ActiveSync status ${folderSync.status} after reprovision retry. Raw response: ${compactXml(
            folderSyncXml,
          )}`,
        );
      }
      await this.performProvisioning();
      return this.resolveMailFolders(reprovisionAttempts + 1);
    }

    if (folderSync.status !== "1") {
      throw new Error(
        `FolderSync failed with ActiveSync status ${folderSync.status}. Raw response: ${compactXml(folderSyncXml)}`,
      );
    }

    const inbox = findInboxFolder(folderSync.folders);
    if (!inbox) {
      throw new Error(
        `Inbox folder not found in FolderSync response. Folders seen: ${JSON.stringify(
          folderSync.folders,
        )}. Raw response: ${compactXml(folderSyncXml)}`,
      );
    }

    return { inbox, sent: findOutboxFolder(folderSync.folders) };
  }
```

Then change `getMailboxMessages` so its FolderSync section reads:

```ts
  async getMailboxMessages({
    days,
    since,
    maxPages = 50,
    windowSize = 50,
  }: MailboxMessageOptions = {}): Promise<NormalizedMailboxMessage[]> {
    void days;

    const { inbox, sent } = await this.resolveMailFolders();

    const inboxMessages = await this.syncFolderMessages({
      mailbox: "inbox",
      collectionId: inbox.serverId,
      maxPages,
      windowSize,
      since,
    });
    const sentMessages = sent
      ? await this.syncFolderMessages({
          mailbox: "sent",
          collectionId: sent.serverId,
          maxPages,
          windowSize,
          since,
        })
      : [];

    return [...inboxMessages, ...sentMessages];
  }
```

Update the `MailboxMessageOptions` interface (top of file) to add `since`:

```ts
interface MailboxMessageOptions {
  days?: number;
  since?: Date;
  maxPages?: number;
  windowSize?: number;
}
```

(Note: `reprovisionAttempts` moved into `resolveMailFolders`, so remove it from `MailboxMessageOptions` and from the `getMailboxMessages` signature.) The `syncFolderMessages` `since` parameter is wired up in Task 6 — for now add `since` to its destructured params and ignore it:

```ts
  async syncFolderMessages({
    mailbox,
    collectionId,
    maxPages,
    windowSize,
    since,
  }: {
    mailbox: MailboxKind;
    collectionId: string;
    maxPages: number;
    windowSize: number;
    since?: Date;
  }): Promise<NormalizedMailboxMessage[]> {
    void since;
    // ... existing body unchanged ...
```

- [ ] **Step 3: Add `fetchMessage` and `getMessageById`**

Add these methods to the class:

```ts
  async fetchMessage({
    collectionId,
    serverId,
  }: {
    collectionId: string;
    serverId: string;
  }): Promise<NormalizedMailboxMessage["message"] | null> {
    const xml = await this.executeCommand(
      "ItemOperations",
      buildItemOperationsFetchXml({
        protocolVersion: this.config.protocolVersion,
        collectionId,
        serverId,
      }),
    );

    const parsed = parseItemOperationsXml(xml);
    if (!parsed.message) {
      return null;
    }

    return normalizeSyncMessages([parsed.message])[0] ?? null;
  }

  async getMessageById({ serverId }: { serverId: string }): Promise<NormalizedMailboxMessage | null> {
    const { inbox, sent } = await this.resolveMailFolders();

    const candidates: Array<{ mailbox: MailboxKind; collectionId: string }> = [
      { mailbox: "inbox", collectionId: inbox.serverId },
    ];
    if (sent) {
      candidates.push({ mailbox: "sent", collectionId: sent.serverId });
    }

    for (const candidate of candidates) {
      const message = await this.fetchMessage({ collectionId: candidate.collectionId, serverId });
      if (message) {
        return { mailbox: candidate.mailbox, message };
      }
    }

    return null;
  }
```

- [ ] **Step 4: Write the test at `src/activesync/client.test.ts`**

This test feeds the client WBXML bytes built with `encodeWbxml`, so the builder, code page, and parser are exercised end-to-end against a mocked transport.

```ts
import { describe, expect, it, vi } from "vitest";

import { ActiveSyncClient } from "./client.js";
import { encodeWbxml } from "./wbxml.js";
import type { ActiveSyncClientConfig } from "../types.js";

function wbxmlResponse(xml: string) {
  const buffer = encodeWbxml(xml);
  return {
    ok: true,
    status: 200,
    arrayBuffer: async () => buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength),
  };
}

const FOLDER_SYNC_XML =
  '<?xml version="1.0" encoding="utf-8"?>' +
  '<FolderSync xmlns="FolderHierarchy:"><Status>1</Status><SyncKey>1</SyncKey><Changes>' +
  "<Add><ServerId>2</ServerId><ParentId>0</ParentId><DisplayName>Inbox</DisplayName><Type>2</Type></Add>" +
  "<Add><ServerId>5</ServerId><ParentId>0</ParentId><DisplayName>Sent Items</DisplayName><Type>5</Type></Add>" +
  "</Changes></FolderSync>";

const ITEM_OPS_XML =
  '<?xml version="1.0" encoding="utf-8"?>' +
  '<ItemOperations xmlns="ItemOperations:" xmlns:airsync="AirSync:" xmlns:airsyncbase="AirSyncBase:">' +
  "<Status>1</Status><Response><Fetch><Status>1</Status>" +
  "<airsync:CollectionId>2</airsync:CollectionId><airsync:ServerId>2:42</airsync:ServerId>" +
  "<Properties><Subject>Hello</Subject><From>a@b.com</From>" +
  "<DateReceived>2026-06-17T09:00:00.000Z</DateReceived>" +
  "<airsyncbase:Body><airsyncbase:Type>1</airsyncbase:Type><airsyncbase:Data>Body text</airsyncbase:Data></airsyncbase:Body>" +
  "</Properties></Fetch></Response></ItemOperations>";

function makeConfig(): ActiveSyncClientConfig {
  return {
    email: "user@example.com",
    username: "user@example.com",
    password: "secret",
    host: "example.com",
    endpoint: "https://example.com/Microsoft-Server-ActiveSync",
    verbose: false,
    deviceId: "dev",
    deviceType: "iPhone",
    userAgent: "agent",
    protocolVersion: "14.1",
    deviceModel: "iPhone",
    deviceImei: "0",
    deviceFriendlyName: "fn",
    deviceOs: "iOS",
    deviceOsLanguage: "en-us",
    devicePhoneNumber: "0",
    deviceMobileOperator: "x",
  };
}

describe("ActiveSyncClient.getMessageById", () => {
  it("resolves folders then fetches the email from the inbox", async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(wbxmlResponse(FOLDER_SYNC_XML)) // FolderSync
      .mockResolvedValueOnce(wbxmlResponse(ITEM_OPS_XML)); // ItemOperations Fetch

    const client = new ActiveSyncClient(makeConfig(), { fetchImpl, logger: { error() {} } });
    const result = await client.getMessageById({ serverId: "2:42" });

    expect(result?.mailbox).toBe("inbox");
    expect(result?.message.subject).toBe("Hello");
    expect(result?.message.bodyText).toBe("Body text");
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });
});
```

- [ ] **Step 5: Run the test**

Run: `yarn test src/activesync/client.test.ts`
Expected: PASS (1 test).

- [ ] **Step 6: Confirm nothing else broke + type-check**

Run: `yarn test && yarn typecheck`
Expected: all PASS, no type errors.

- [ ] **Step 7: Commit**

```bash
git add src/activesync/client.ts src/activesync/client.test.ts
git commit -m "feat: fetch a single email by id via ItemOperations"
```

---

## Task 6: Cutoff-aware paging + truncation warning

Honor the requested time window when paging (spec §6.1): stop when the server has no more data, when an entire page is older than the cutoff, or at the `maxPages` safety cap — and warn on stderr if the cap truncates results.

**Files:**
- Modify: `src/activesync/client.ts`
- Test: `src/activesync/paging.test.ts`

- [ ] **Step 1: Implement the `since`-aware stop conditions in `syncFolderMessages`**

Replace the `void since;` line and the tail of the loop. The full updated method body:

```ts
  async syncFolderMessages({
    mailbox,
    collectionId,
    maxPages,
    windowSize,
    since,
  }: {
    mailbox: MailboxKind;
    collectionId: string;
    maxPages: number;
    windowSize: number;
    since?: Date;
  }): Promise<NormalizedMailboxMessage[]> {
    const messages: NormalizedMailboxMessage[] = [];
    let syncKey = "0";

    for (let page = 0; page < maxPages; page += 1) {
      const requestSyncKey = syncKey;
      const xml = await this.executeCommand(
        "Sync",
        buildSyncRequestXml({
          protocolVersion: this.config.protocolVersion,
          syncKey,
          collectionId,
          windowSize,
        }),
      );
      const parsed = parseSyncXml(xml);
      if (parsed.status && parsed.status !== "1") {
        throw new Error(`Sync failed with ActiveSync status ${parsed.status}. Raw response: ${compactXml(xml)}`);
      }
      syncKey = parsed.syncKey || syncKey;

      const pageMessages = normalizeSyncMessages(parsed.messages).map((message) => ({
        mailbox,
        message,
      }));
      messages.push(...pageMessages);

      if (requestSyncKey === "0" && parsed.messages.length === 0 && syncKey !== "0") {
        continue;
      }

      if (
        since &&
        pageMessages.length > 0 &&
        pageMessages.every((entry) => isOlderThan(entry.message.receivedAt, since))
      ) {
        break;
      }

      if (!parsed.moreAvailable) {
        break;
      }

      if (page === maxPages - 1) {
        this.logger.error(
          `Warning: reached the ${maxPages}-page sync limit for ${mailbox}; messages within the requested window may be missing. Increase --max-pages.`,
        );
      }
    }

    return messages;
  }
```

Add this helper near `compactXml` at the bottom of the file:

```ts
function isOlderThan(receivedAt: string, cutoff: Date): boolean {
  const date = new Date(receivedAt);
  return !Number.isNaN(date.valueOf()) && date < cutoff;
}
```

- [ ] **Step 2: Write the test at `src/activesync/paging.test.ts`**

```ts
import { describe, expect, it, vi } from "vitest";

import { ActiveSyncClient } from "./client.js";
import { encodeWbxml } from "./wbxml.js";
import type { ActiveSyncClientConfig } from "../types.js";

function wbxmlResponse(xml: string) {
  const buffer = encodeWbxml(xml);
  return {
    ok: true,
    status: 200,
    arrayBuffer: async () => buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength),
  };
}

function syncPage({ syncKey, receivedAt, more }: { syncKey: string; receivedAt: string; more: boolean }) {
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<Sync xmlns="AirSync:"><Collections><Collection><SyncKey>' +
    syncKey +
    "</SyncKey><Status>1</Status>" +
    (more ? "<MoreAvailable/>" : "") +
    "<Commands><Add><ServerId>2:" +
    syncKey +
    "</ServerId><ApplicationData><Subject>S</Subject><From>a@b.com</From><DateReceived>" +
    receivedAt +
    "</DateReceived><Type>1</Type><Data>B</Data></ApplicationData></Add></Commands>" +
    "</Collection></Collections></Sync>"
  );
}

function makeConfig(): ActiveSyncClientConfig {
  return {
    email: "user@example.com",
    username: "user@example.com",
    password: "secret",
    host: "example.com",
    endpoint: "https://example.com/Microsoft-Server-ActiveSync",
    verbose: false,
    deviceId: "dev",
    deviceType: "iPhone",
    userAgent: "agent",
    protocolVersion: "14.1",
    deviceModel: "iPhone",
    deviceImei: "0",
    deviceFriendlyName: "fn",
    deviceOs: "iOS",
    deviceOsLanguage: "en-us",
    devicePhoneNumber: "0",
    deviceMobileOperator: "x",
  };
}

describe("syncFolderMessages cutoff paging", () => {
  it("stops paging once a whole page is older than the cutoff", async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(wbxmlResponse(syncPage({ syncKey: "1", receivedAt: "2026-06-17T11:00:00.000Z", more: true })))
      .mockResolvedValueOnce(wbxmlResponse(syncPage({ syncKey: "2", receivedAt: "2026-06-10T11:00:00.000Z", more: true })));

    const client = new ActiveSyncClient(makeConfig(), { fetchImpl, logger: { error() {} } });
    const since = new Date("2026-06-17T00:00:00.000Z");

    const result = await client.syncFolderMessages({
      mailbox: "inbox",
      collectionId: "2",
      maxPages: 50,
      windowSize: 50,
      since,
    });

    // Page 1 (in-window) + page 2 (older, triggers stop) fetched; no third page.
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    expect(result).toHaveLength(2);
  });
});
```

- [ ] **Step 3: Run the test**

Run: `yarn test src/activesync/paging.test.ts`
Expected: PASS (1 test).

- [ ] **Step 4: Commit**

```bash
git add src/activesync/client.ts src/activesync/paging.test.ts
git commit -m "feat: page sync until the requested time window is covered"
```

---

## Task 7: Config (file + keyring)

Adapt `cli-example/config.ts` for the `outlook` service. Stores `email`, `username`, `host`, optional `endpoint`/`protocolVersion` in `~/.config/outlook/config.json`; password in keyring service `outlook`, account `cli`.

**Files:**
- Create: `src/config.ts`
- Test: `src/config.test.ts`

- [ ] **Step 1: Write `src/config.ts`**

```ts
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { homedir } from "node:os";
import { dirname, join } from "node:path";

const KEYRING_SERVICE_NAME = "outlook";
const KEYRING_ACCOUNT_NAME = "cli";

export function getConfigPath(home = homedir()): string {
  return join(home, ".config", "outlook");
}

export function getFileConfigPath(home = homedir()): string {
  return join(getConfigPath(home), "config.json");
}

export type FileConfig = {
  email?: string;
  username?: string;
  host?: string;
  endpoint?: string;
  protocolVersion?: string;
};

export type ResolvedConfig = {
  email: string;
  username: string;
  host: string;
  password: string;
  endpoint?: string;
  protocolVersion?: string;
};

export async function loadConfig(): Promise<ResolvedConfig> {
  const { email, username, host, endpoint, protocolVersion } = await readConfigFile(getFileConfigPath());
  const password = await readPassword();

  if (!email || !username || !host || !password) {
    throw new Error("Missing required configuration: email, username, host or password. Run `outlook config init`.");
  }

  return { email, username, host, password, endpoint, protocolVersion };
}

async function readConfigFile(filePath: string): Promise<FileConfig> {
  try {
    return JSON.parse(await readFile(filePath, "utf8")) as FileConfig;
  } catch (error) {
    if (isMissingFile(error)) {
      return {};
    }
    throw error;
  }
}

export async function writeConfig(config: ResolvedConfig): Promise<void> {
  const configPath = getFileConfigPath();
  const { password, ...fileConfig } = config;

  await mkdir(dirname(configPath), { recursive: true });
  await writeFile(configPath, `${JSON.stringify(fileConfig, null, 2)}\n`, "utf8");
  await savePassword(password);
}

async function getKeytar() {
  try {
    const { default: keytar } = await import("keytar");
    return keytar;
  } catch {
    throw new Error(
      "keytar is not available. Install libsecret (libsecret-1-dev on Debian/Ubuntu, libsecret-devel on RHEL/Fedora) and reinstall.",
    );
  }
}

async function readPassword() {
  const keytar = await getKeytar();
  const results = await keytar.findCredentials(KEYRING_SERVICE_NAME);
  if (!results || results.length === 0) {
    return null;
  }
  return results.find((cred) => cred.account === KEYRING_ACCOUNT_NAME)?.password ?? null;
}

export async function savePassword(password: string) {
  const keytar = await getKeytar();
  await keytar.setPassword(KEYRING_SERVICE_NAME, KEYRING_ACCOUNT_NAME, password);
}

export function redactConfig(config: ResolvedConfig): ResolvedConfig {
  return { ...config, password: "****" };
}

function isMissingFile(error: unknown): error is NodeJS.ErrnoException {
  return typeof error === "object" && error !== null && "code" in error && error.code === "ENOENT";
}
```

- [ ] **Step 2: Write `src/config.test.ts`**

```ts
import { beforeEach, describe, expect, it, vi } from "vitest";

import { getConfigPath, getFileConfigPath, loadConfig, redactConfig, writeConfig } from "./config.js";

const { mockFindCredentials, mockSetPassword, mockMkdir, mockReadFile, mockWriteFile } = vi.hoisted(() => ({
  mockFindCredentials: vi.fn(),
  mockSetPassword: vi.fn(),
  mockMkdir: vi.fn(),
  mockReadFile: vi.fn(),
  mockWriteFile: vi.fn(),
}));

vi.mock("keytar", () => ({
  default: { findCredentials: mockFindCredentials, setPassword: mockSetPassword },
}));

vi.mock("node:fs/promises", () => ({
  mkdir: mockMkdir,
  readFile: mockReadFile,
  writeFile: mockWriteFile,
}));

describe("config", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("builds config paths under .config/outlook", () => {
    expect(getConfigPath("/tmp/home")).toBe("/tmp/home/.config/outlook");
    expect(getFileConfigPath("/tmp/home")).toBe("/tmp/home/.config/outlook/config.json");
  });

  it("loads config from file and keyring", async () => {
    mockReadFile.mockResolvedValueOnce(
      JSON.stringify({ email: "u@e.com", username: "u@e.com", host: "e.com" }),
    );
    mockFindCredentials.mockResolvedValueOnce([
      { account: "other", password: "ignore" },
      { account: "cli", password: "secret" },
    ]);

    await expect(loadConfig()).resolves.toEqual({
      email: "u@e.com",
      username: "u@e.com",
      host: "e.com",
      password: "secret",
      endpoint: undefined,
      protocolVersion: undefined,
    });
  });

  it("throws when required values are missing", async () => {
    mockReadFile.mockRejectedValueOnce(Object.assign(new Error("missing"), { code: "ENOENT" }));
    mockFindCredentials.mockResolvedValueOnce([]);
    await expect(loadConfig()).rejects.toThrow("Missing required configuration");
  });

  it("writes file config without the password and stores password in keyring", async () => {
    await writeConfig({ email: "u@e.com", username: "u@e.com", host: "e.com", password: "secret" });

    expect(mockWriteFile).toHaveBeenCalledWith(
      expect.stringContaining(".config/outlook/config.json"),
      `${JSON.stringify({ email: "u@e.com", username: "u@e.com", host: "e.com" }, null, 2)}\n`,
      "utf8",
    );
    expect(mockSetPassword).toHaveBeenCalledWith("outlook", "cli", "secret");
  });

  it("redacts the password", () => {
    expect(
      redactConfig({ email: "u@e.com", username: "u@e.com", host: "e.com", password: "secret" }),
    ).toEqual({ email: "u@e.com", username: "u@e.com", host: "e.com", password: "****" });
  });
});
```

- [ ] **Step 3: Run the test**

Run: `yarn test src/config.test.ts`
Expected: PASS (5 tests).

- [ ] **Step 4: Commit**

```bash
git add src/config.ts src/config.test.ts
git commit -m "feat: add config storage with keyring-backed password"
```

---

## Task 8: IO helpers (output, prompt, errors)

Adapt `cli-example/io.ts`, dropping the AppSec-specific table formatters. Output modes are `text` (default) and `json`.

**Files:**
- Create: `src/io.ts`
- Test: `src/io.test.ts`

- [ ] **Step 1: Write `src/io.ts`**

```ts
import readline from "node:readline";
import { Writable } from "node:stream";

export type OutputFormat = "text" | "json";

type PrintOutputParams = {
  value: unknown;
  format?: OutputFormat;
  textFormatter?: (value: unknown) => string;
};

export function printOutput({ value, format = "text", textFormatter }: PrintOutputParams): void {
  if (format === "json") {
    process.stdout.write(`${JSON.stringify(value, null, 2)}\n`);
    return;
  }

  if (textFormatter) {
    process.stdout.write(textFormatter(value));
    return;
  }

  process.stdout.write(`${formatLine(value)}\n`);
}

export function formatLine(value: unknown): string {
  return typeof value === "string" ? value : JSON.stringify(value, null, 2);
}

export function parseOutputFormat(value: string | undefined): OutputFormat {
  if (value === undefined) {
    return "text";
  }
  if (value !== "text" && value !== "json") {
    throw new Error(`Unsupported output format: ${value}. Use "text" or "json".`);
  }
  return value;
}

export function formatError(error: unknown): string {
  if (!(error instanceof Error)) {
    return `${String(error)}\n`;
  }

  type ErrorWithDetails = Error & { cause?: unknown; code?: unknown };
  const messages = [error.message];
  let { cause } = error as ErrorWithDetails;

  while (cause) {
    if (cause instanceof Error) {
      const code = typeof (cause as ErrorWithDetails).code === "string" ? ` (${(cause as ErrorWithDetails).code})` : "";
      messages.push(`${cause.message}${code}`);
      cause = (cause as ErrorWithDetails).cause;
    } else {
      messages.push(String(cause));
      cause = undefined;
    }
  }

  return `${messages.join("\nCaused by: ")}\n`;
}

export class MutableStdout extends Writable {
  public muted = false;

  // eslint-disable-next-line no-underscore-dangle, class-methods-use-this
  _write(chunk: Buffer, _encoding: unknown, callback: (error?: Error | null) => void) {
    if (this.muted) {
      const str = chunk.toString();
      // eslint-disable-next-line no-control-regex
      if (/^(?:\x1B|\r|\n|\x08|\x7f)/.test(str)) {
        process.stdout.write(chunk);
      } else {
        process.stdout.write("*".repeat(str.length));
      }
    } else {
      process.stdout.write(chunk);
    }
    callback();
  }
}

type PromptParams = {
  message: string;
  defaultValue?: string;
  secret?: boolean;
};

type PromptDependencies = {
  createInterface?: typeof readline.createInterface;
};

export async function prompt(
  { message, defaultValue = "", secret = false }: PromptParams,
  dependencies: PromptDependencies = {},
): Promise<string> {
  const outputStream = new MutableStdout();
  const createInterface = dependencies.createInterface ?? readline.createInterface;
  const rl = createInterface({ input: process.stdin, output: outputStream, terminal: true });

  try {
    return await new Promise<string>((resolve) => {
      rl.question(
        `${message.trim()}${defaultValue ? ` (leave empty for default, ${defaultValue})` : ""}: `,
        (answer: string) => {
          outputStream.muted = false;
          if (secret) {
            process.stdout.write("\n");
          }
          resolve(answer || defaultValue);
        },
      );
      if (secret) {
        outputStream.muted = true;
      }
    });
  } finally {
    rl.close();
  }
}
```

- [ ] **Step 2: Write `src/io.test.ts`**

```ts
import { afterEach, describe, expect, it, vi } from "vitest";

import { formatError, formatLine, parseOutputFormat, printOutput } from "./io.js";

describe("parseOutputFormat", () => {
  it("defaults to text and rejects unknown values", () => {
    expect(parseOutputFormat(undefined)).toBe("text");
    expect(parseOutputFormat("json")).toBe("json");
    expect(() => parseOutputFormat("xml")).toThrow("Unsupported output format");
  });
});

describe("formatLine", () => {
  it("returns strings unchanged and stringifies objects", () => {
    expect(formatLine("hi")).toBe("hi");
    expect(formatLine({ a: 1 })).toBe(JSON.stringify({ a: 1 }, null, 2));
  });
});

describe("formatError", () => {
  it("chains causes", () => {
    const error = new Error("top", { cause: new Error("inner") });
    expect(formatError(error)).toBe("top\nCaused by: inner\n");
  });
});

describe("printOutput", () => {
  afterEach(() => vi.restoreAllMocks());

  it("writes pretty JSON in json mode", () => {
    const write = vi.spyOn(process.stdout, "write").mockReturnValue(true);
    printOutput({ value: { a: 1 }, format: "json" });
    expect(write).toHaveBeenCalledWith(`${JSON.stringify({ a: 1 }, null, 2)}\n`);
  });

  it("uses the text formatter in text mode", () => {
    const write = vi.spyOn(process.stdout, "write").mockReturnValue(true);
    printOutput({ value: ["x"], format: "text", textFormatter: () => "FORMATTED" });
    expect(write).toHaveBeenCalledWith("FORMATTED");
  });
});
```

- [ ] **Step 3: Run the test**

Run: `yarn test src/io.test.ts`
Expected: PASS (5 tests).

- [ ] **Step 4: Commit**

```bash
git add src/io.ts src/io.test.ts
git commit -m "feat: add IO helpers for output, prompts, and error formatting"
```

---

## Task 9: Plain-text email formatters

**Files:**
- Create: `src/format.ts`
- Test: `src/format.test.ts`

- [ ] **Step 1: Write `src/format.ts`**

```ts
import type { NormalizedMailboxMessage } from "./types.js";

const SEP = " · ";

export function formatEmailList(
  entries: NormalizedMailboxMessage[],
  options: { includeBody?: boolean } = {},
): string {
  if (entries.length === 0) {
    return "No emails found in the requested window.\n";
  }

  if (!options.includeBody) {
    return `${entries.map(formatListRow).join("\n")}\n`;
  }

  return `${entries.map((entry) => `${formatListRow(entry)}\n${resolveBody(entry)}\n`).join("\n")}`;
}

export function formatEmail(entry: NormalizedMailboxMessage): string {
  const { message } = entry;
  const lines = [
    `Mailbox: ${entry.mailbox}`,
    `Id:      ${message.id ?? ""}`,
    `From:    ${message.from || "Unknown sender"}`,
    message.to.length > 0 ? `To:      ${message.to.join(", ")}` : null,
    message.cc.length > 0 ? `Cc:      ${message.cc.join(", ")}` : null,
    `Received: ${message.receivedAt}`,
    `Subject: ${message.subject || "(no subject)"}`,
    "",
    resolveBody(entry),
  ].filter((line): line is string => line !== null);

  return `${lines.join("\n")}\n`;
}

function formatListRow(entry: NormalizedMailboxMessage): string {
  const { message } = entry;
  return [
    entry.mailbox,
    message.receivedAt || "(no date)",
    message.id ?? "",
    message.subject || "(no subject)",
    message.from || "(unknown)",
  ].join(SEP);
}

function resolveBody(entry: NormalizedMailboxMessage): string {
  const bodyText = entry.message.bodyText.trim();
  if (bodyText) {
    return bodyText;
  }
  const preview = entry.message.preview.trim();
  return preview || "(no body)";
}
```

- [ ] **Step 2: Write `src/format.test.ts`**

```ts
import { describe, expect, it } from "vitest";

import { formatEmail, formatEmailList } from "./format.js";
import type { NormalizedMailboxMessage } from "./types.js";

function entry(overrides: Partial<NormalizedMailboxMessage["message"]> = {}): NormalizedMailboxMessage {
  return {
    mailbox: "inbox",
    message: {
      id: "2:1",
      subject: "Subject",
      from: "a@b.com",
      to: ["c@d.com"],
      cc: [],
      receivedAt: "2026-06-17T09:00:00.000Z",
      preview: "preview",
      bodyText: "Full body",
      bodyHtml: "",
      ...overrides,
    },
  };
}

describe("formatEmailList", () => {
  it("renders a title-only row with the · separator", () => {
    const out = formatEmailList([entry()]);
    expect(out).toBe("inbox · 2026-06-17T09:00:00.000Z · 2:1 · Subject · a@b.com\n");
  });

  it("includes the body when requested", () => {
    const out = formatEmailList([entry()], { includeBody: true });
    expect(out).toContain("Full body");
  });

  it("reports an empty result", () => {
    expect(formatEmailList([])).toBe("No emails found in the requested window.\n");
  });
});

describe("formatEmail", () => {
  it("renders headers and falls back to preview when bodyText is empty", () => {
    const out = formatEmail(entry({ bodyText: "" }));
    expect(out).toContain("Subject: Subject");
    expect(out).toContain("preview");
  });
});
```

- [ ] **Step 3: Run the test**

Run: `yarn test src/format.test.ts`
Expected: PASS (4 tests).

- [ ] **Step 4: Commit**

```bash
git add src/format.ts src/format.test.ts
git commit -m "feat: add plain-text email formatters"
```

---

## Task 10: OutlookClient wrapper

Wraps `ActiveSyncClient` with device defaults, hours filtering, and mailbox selection. Uses dependency injection (`createClient`) for testing, mirroring `outlook-example/client.ts`.

**Files:**
- Create: `src/client.ts`
- Test: `src/client.test.ts`

- [ ] **Step 1: Write `src/client.ts`**

```ts
import { ActiveSyncClient } from "./activesync/client.js";
import type { ResolvedConfig } from "./config.js";
import { filterMessagesWithinHours } from "./messages.js";
import type { ActiveSyncClientConfig, MailboxKind, NormalizedMailboxMessage } from "./types.js";

const DEVICE_DEFAULTS = {
  verbose: false,
  endpoint: "",
  deviceId: "NodeiPhoneClient001",
  deviceType: "iPhone",
  userAgent: "Apple-iPhone14C3/1704.10",
  protocolVersion: "14.1",
  deviceModel: "iPhone",
  deviceImei: "000000000000000",
  deviceFriendlyName: "Outlook CLI",
  deviceOs: "iOS 18.0",
  deviceOsLanguage: "en-us",
  devicePhoneNumber: "0000000000",
  deviceMobileOperator: "Unknown",
} as const;

export interface ActiveSyncLike {
  resolveMailFolders(): Promise<unknown>;
  getMailboxMessages(opts: { since?: Date; maxPages?: number }): Promise<NormalizedMailboxMessage[]>;
  getMessageById(opts: { serverId: string }): Promise<NormalizedMailboxMessage | null>;
}

interface OutlookClientDependencies {
  createClient?: (config: ActiveSyncClientConfig) => ActiveSyncLike;
}

export function buildClientConfig(config: ResolvedConfig): ActiveSyncClientConfig {
  return {
    ...DEVICE_DEFAULTS,
    email: config.email,
    username: config.username,
    password: config.password,
    host: config.host,
    endpoint: config.endpoint ?? DEVICE_DEFAULTS.endpoint,
    protocolVersion: config.protocolVersion ?? DEVICE_DEFAULTS.protocolVersion,
  };
}

export class OutlookClient {
  private createClient: (config: ActiveSyncClientConfig) => ActiveSyncLike;

  constructor(
    private config: ResolvedConfig,
    dependencies: OutlookClientDependencies = {},
  ) {
    this.createClient =
      dependencies.createClient ?? ((clientConfig) => new ActiveSyncClient(clientConfig, { logger: console }));
  }

  private build(): ActiveSyncLike {
    return this.createClient(buildClientConfig(this.config));
  }

  async verifyConnection(): Promise<void> {
    await this.build().resolveMailFolders();
  }

  async listEmails({
    hours,
    mailbox,
    maxPages,
  }: {
    hours: number;
    mailbox?: MailboxKind;
    maxPages?: number;
  }): Promise<NormalizedMailboxMessage[]> {
    const since = new Date(Date.now() - hours * 60 * 60 * 1000);
    const all = await this.build().getMailboxMessages({ since, maxPages });

    const withinWindow = filterMessagesWithinHours(all, hours);
    const scoped = mailbox ? withinWindow.filter((entry) => entry.mailbox === mailbox) : withinWindow;

    return scoped.sort(
      (left, right) => new Date(right.message.receivedAt).valueOf() - new Date(left.message.receivedAt).valueOf(),
    );
  }

  async getEmail({ id }: { id: string }): Promise<NormalizedMailboxMessage | null> {
    return this.build().getMessageById({ serverId: id });
  }
}
```

- [ ] **Step 2: Write `src/client.test.ts`**

```ts
import { describe, expect, it, vi } from "vitest";

import { OutlookClient, buildClientConfig } from "./client.js";
import type { ResolvedConfig } from "./config.js";
import type { NormalizedMailboxMessage } from "./types.js";

const config: ResolvedConfig = {
  email: "u@e.com",
  username: "u@e.com",
  host: "e.com",
  password: "secret",
};

function message(mailbox: "inbox" | "sent", receivedAt: string): NormalizedMailboxMessage {
  return {
    mailbox,
    message: {
      id: receivedAt,
      subject: "s",
      from: "f",
      to: [],
      cc: [],
      receivedAt,
      preview: "",
      bodyText: "",
      bodyHtml: "",
    },
  };
}

describe("buildClientConfig", () => {
  it("merges device defaults with user config", () => {
    const built = buildClientConfig(config);
    expect(built.email).toBe("u@e.com");
    expect(built.deviceType).toBe("iPhone");
    expect(built.protocolVersion).toBe("14.1");
  });
});

describe("OutlookClient.listEmails", () => {
  it("filters by mailbox and sorts newest-first", async () => {
    const now = Date.now();
    const recent = new Date(now - 60 * 60 * 1000).toISOString();
    const older = new Date(now - 2 * 60 * 60 * 1000).toISOString();
    const stub = {
      resolveMailFolders: vi.fn(),
      getMailboxMessages: vi.fn().mockResolvedValue([
        message("inbox", older),
        message("sent", recent),
        message("inbox", recent),
      ]),
      getMessageById: vi.fn(),
    };

    const client = new OutlookClient(config, { createClient: () => stub });
    const result = await client.listEmails({ hours: 24, mailbox: "inbox" });

    expect(result.map((e) => e.mailbox)).toEqual(["inbox", "inbox"]);
    expect(result[0].message.receivedAt).toBe(recent);
  });
});

describe("OutlookClient.getEmail", () => {
  it("delegates to getMessageById", async () => {
    const found = message("inbox", new Date().toISOString());
    const stub = {
      resolveMailFolders: vi.fn(),
      getMailboxMessages: vi.fn(),
      getMessageById: vi.fn().mockResolvedValue(found),
    };
    const client = new OutlookClient(config, { createClient: () => stub });
    await expect(client.getEmail({ id: "2:1" })).resolves.toBe(found);
    expect(stub.getMessageById).toHaveBeenCalledWith({ serverId: "2:1" });
  });
});
```

- [ ] **Step 3: Run the test**

Run: `yarn test src/client.test.ts`
Expected: PASS (3 tests).

- [ ] **Step 4: Commit**

```bash
git add src/client.ts src/client.test.ts
git commit -m "feat: add OutlookClient wrapper with hours filtering and mailbox scoping"
```

---

## Task 11: CLI wiring + entry point

**Files:**
- Create: `src/cli.ts`, `src/index.ts`
- Test: `src/cli.test.ts`

- [ ] **Step 1: Write `src/cli.ts`**

```ts
import { Command } from "commander";

import { OutlookClient } from "./client.js";
import { getFileConfigPath, loadConfig, redactConfig, writeConfig, type ResolvedConfig } from "./config.js";
import { formatEmail, formatEmailList } from "./format.js";
import { formatError, parseOutputFormat, printOutput, prompt } from "./io.js";
import type { MailboxKind, NormalizedMailboxMessage } from "./types.js";

function parseMailbox(value: string | undefined): MailboxKind | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (value !== "inbox" && value !== "sent") {
    throw new Error(`Unsupported mailbox: ${value}. Use "inbox" or "sent".`);
  }
  return value;
}

function parsePositiveInt(value: string | undefined, fallback: number, label: string): number {
  if (value === undefined) {
    return fallback;
  }
  const parsed = Number.parseInt(value, 10);
  if (Number.isNaN(parsed) || parsed <= 0) {
    throw new Error(`${label} must be a positive integer, got "${value}".`);
  }
  return parsed;
}

export function createCli() {
  const program = new Command();
  program.name("outlook").description("Read Outlook/Exchange mail over ActiveSync from the command line.");

  const configCommand = program.command("config").description("Create, verify, and inspect the connection settings.");

  configCommand
    .command("init")
    .description("Prompt for connection settings, verify them, and save them locally.")
    .action(async () => {
      const email = await prompt({ message: "Email" });
      const username = await prompt({ message: "Username", defaultValue: email });
      const host = await prompt({ message: "Host (e.g. mail.example.com)" });
      const password = await prompt({ message: "Password", secret: true });

      const config: ResolvedConfig = { email, username, host, password };
      await new OutlookClient(config).verifyConnection();
      await writeConfig(config);

      process.stdout.write(`Verified and saved config to ${getFileConfigPath()}\n`);
    });

  configCommand
    .command("verify")
    .description("Verify the saved configuration by connecting to the server.")
    .action(async () => {
      await new OutlookClient(await loadConfig()).verifyConnection();
      process.stdout.write("Connection verified.\n");
    });

  configCommand
    .command("show")
    .description("Print the saved configuration with the password redacted.")
    .action(async () => {
      printOutput({ value: redactConfig(await loadConfig()), format: "json" });
    });

  const emailsCommand = program.command("emails").description("List and fetch emails.");

  emailsCommand
    .command("list")
    .description("List emails received in the last N hours.")
    .option("--hours <N>", "Lookback window in hours.", "24")
    .option("--mailbox <inbox|sent>", "Restrict to one mailbox. Defaults to both.")
    .option("--body", "Include the full plain-text body for each email.")
    .option("--max-pages <N>", "Safety cap on sync pages per mailbox.", "50")
    .option("--output <text|json>", "Output format: text or json.", "text")
    .action(
      async (options: { hours?: string; mailbox?: string; body?: boolean; maxPages?: string; output?: string }) => {
        const format = parseOutputFormat(options.output);
        const hours = parsePositiveInt(options.hours, 24, "--hours");
        const maxPages = parsePositiveInt(options.maxPages, 50, "--max-pages");
        const mailbox = parseMailbox(options.mailbox);

        const client = new OutlookClient(await loadConfig());
        const emails = await client.listEmails({ hours, mailbox, maxPages });

        printOutput({
          value: emails,
          format,
          textFormatter: (value) =>
            formatEmailList(value as NormalizedMailboxMessage[], { includeBody: Boolean(options.body) }),
        });
      },
    );

  emailsCommand
    .command("get <id>")
    .description("Fetch a single full email by id.")
    .option("--output <text|json>", "Output format: text or json.", "text")
    .action(async (id: string, options: { output?: string }) => {
      const format = parseOutputFormat(options.output);
      const client = new OutlookClient(await loadConfig());
      const email = await client.getEmail({ id });

      if (!email) {
        throw new Error(`Email with id "${id}" was not found in the Inbox or Sent folder.`);
      }

      printOutput({
        value: email,
        format,
        textFormatter: (value) => formatEmail(value as NormalizedMailboxMessage),
      });
    });

  return {
    program,
    async run(argv: string[]): Promise<number> {
      try {
        await program.parseAsync(argv, { from: "user" });
        return 0;
      } catch (error) {
        process.stderr.write(formatError(error));
        return 1;
      }
    },
  };
}
```

- [ ] **Step 2: Write `src/index.ts`**

```ts
#!/usr/bin/env node

import { createCli } from "./cli.js";

async function main() {
  const cli = createCli();
  const exitCode = await cli.run(process.argv.slice(2));
  if (exitCode !== 0) {
    process.exitCode = exitCode;
  }
}

main();
```

- [ ] **Step 3: Write `src/cli.test.ts`**

Mocks `config` and `client` so the CLI is tested without network or keyring.

```ts
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const { mockLoadConfig, mockListEmails, mockGetEmail, mockVerify } = vi.hoisted(() => ({
  mockLoadConfig: vi.fn(),
  mockListEmails: vi.fn(),
  mockGetEmail: vi.fn(),
  mockVerify: vi.fn(),
}));

vi.mock("./config.js", () => ({
  loadConfig: mockLoadConfig,
  getFileConfigPath: () => "/tmp/config.json",
  redactConfig: (c: unknown) => c,
  writeConfig: vi.fn(),
}));

vi.mock("./client.js", () => ({
  OutlookClient: vi.fn().mockImplementation(() => ({
    listEmails: mockListEmails,
    getEmail: mockGetEmail,
    verifyConnection: mockVerify,
  })),
}));

import { createCli } from "./cli.js";

describe("createCli", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockLoadConfig.mockResolvedValue({ email: "u@e.com", username: "u@e.com", host: "e.com", password: "p" });
    vi.spyOn(process.stdout, "write").mockReturnValue(true);
    vi.spyOn(process.stderr, "write").mockReturnValue(true);
  });

  afterEach(() => vi.restoreAllMocks());

  it("lists emails and returns exit code 0", async () => {
    mockListEmails.mockResolvedValue([]);
    const exitCode = await createCli().run(["emails", "list", "--hours", "12"]);
    expect(exitCode).toBe(0);
    expect(mockListEmails).toHaveBeenCalledWith({ hours: 12, mailbox: undefined, maxPages: 50 });
  });

  it("returns exit code 1 when get finds no email", async () => {
    mockGetEmail.mockResolvedValue(null);
    const exitCode = await createCli().run(["emails", "get", "2:1"]);
    expect(exitCode).toBe(1);
    expect(process.stderr.write).toHaveBeenCalledWith(expect.stringContaining("was not found"));
  });

  it("rejects an invalid output format with exit code 1", async () => {
    const exitCode = await createCli().run(["emails", "list", "--output", "xml"]);
    expect(exitCode).toBe(1);
    expect(process.stderr.write).toHaveBeenCalledWith(expect.stringContaining("Unsupported output format"));
  });
});
```

- [ ] **Step 4: Run the test**

Run: `yarn test src/cli.test.ts`
Expected: PASS (3 tests).

- [ ] **Step 5: Full suite + type-check + build**

Run: `yarn test && yarn typecheck && yarn build`
Expected: all tests PASS, no type errors, `dist/` produced with `dist/index.js`.

- [ ] **Step 6: Commit**

```bash
git add src/cli.ts src/index.ts src/cli.test.ts
git commit -m "feat: wire up commander CLI and entry point"
```

---

## Task 12: README + manual smoke test

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Replace `README.md` with usage docs**

```markdown
# outlook

A command-line client for reading Outlook/Exchange mailboxes over ActiveSync.

## Install

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
```

- [ ] **Step 2: Manual smoke test (requires real credentials — run if available)**

```bash
node dist/index.js config init
node dist/index.js emails list --hours 24
# copy an id from the list output:
node dist/index.js emails get <id>
node dist/index.js emails list --body --output json
```

Expected: `config init` reports "Verified and saved"; `list` prints rows; `get` prints one full email. If you have no server handy, skip this step and rely on the unit suite.

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: document outlook CLI usage"
```

---

## Self-review notes (already reconciled)

- **Spec coverage:** config init/verify/show (Task 7, 11); list last-N-hours titles (Task 9, 10, 11); list with `--body` (Task 9, 11); get-by-id via ItemOperations (Task 3–5, 11); text/json output (Task 8, 11); Inbox+Sent with `--mailbox` (Task 10, 11); sent/outbox normalization (Task 2); cutoff-aware paging + `--max-pages` warning (Task 6); device defaults (Task 10). All spec sections map to a task.
- **Type consistency:** `ResolvedConfig` (email/username/host/password/endpoint?/protocolVersion?) is defined in Task 7 and consumed unchanged in Tasks 10–11. `ActiveSyncClient` methods `resolveMailFolders`, `getMailboxMessages({since,maxPages})`, `getMessageById`, `fetchMessage` (Tasks 5–6) match the `ActiveSyncLike` interface and stubs in Task 10. `OutlookClient.listEmails/getEmail/verifyConnection` (Task 10) match the CLI calls (Task 11). `printOutput({value,format,textFormatter})` and `parseOutputFormat` (Task 8) match all call sites. `MailboxKind = "inbox" | "sent"` (Task 2) is used consistently in Tasks 9–11.
- **No placeholders:** every code step contains complete, runnable code; every test step contains real assertions; every run step states the command and expected result.
```
