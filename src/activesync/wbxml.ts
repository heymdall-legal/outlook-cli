// @ts-nocheck

const SWITCH_PAGE = 0x00;
const END = 0x01;
const STR_I = 0x03;
const OPAQUE = 0xc3;
const TAG_CONTENT = 0x40;

const CODE_PAGES = [
  createCodePage("AirSync", {
    Sync: 0x05,
    Responses: 0x06,
    Add: 0x07,
    Change: 0x08,
    Delete: 0x09,
    Fetch: 0x0a,
    SyncKey: 0x0b,
    ClientId: 0x0c,
    ServerId: 0x0d,
    Status: 0x0e,
    Collection: 0x0f,
    Class: 0x10,
    CollectionId: 0x12,
    GetChanges: 0x13,
    MoreAvailable: 0x14,
    WindowSize: 0x15,
    Commands: 0x16,
    Options: 0x17,
    FilterType: 0x18,
    Conflict: 0x1b,
    Collections: 0x1c,
    ApplicationData: 0x1d,
    DeletesAsMoves: 0x1e,
    Supported: 0x20,
    SoftDelete: 0x21,
    MIMESupport: 0x22,
    MIMETruncation: 0x23,
    Wait: 0x24,
    Limit: 0x25,
    Partial: 0x26,
    ConversationMode: 0x27,
    MaxItems: 0x28,
    HeartbeatInterval: 0x29,
  }),
  createCodePage("Contacts", {}),
  createCodePage("Email", {
    Attachment: 0x05,
    Attachments: 0x06,
    AttName: 0x07,
    AttSize: 0x08,
    Att0id: 0x09,
    AttMethod: 0x0a,
    Body: 0x0c,
    BodySize: 0x0d,
    BodyTruncated: 0x0e,
    DateReceived: 0x0f,
    DisplayName: 0x10,
    DisplayTo: 0x11,
    Importance: 0x12,
    MessageClass: 0x13,
    Subject: 0x14,
    Read: 0x15,
    To: 0x16,
    Cc: 0x17,
    From: 0x18,
    ReplyTo: 0x19,
    AllDayEvent: 0x1a,
    Categories: 0x1b,
    Category: 0x1c,
    DtStamp: 0x1d,
    EndTime: 0x1e,
    InstanceType: 0x1f,
    BusyStatus: 0x20,
    Location: 0x21,
    MeetingRequest: 0x22,
    Organizer: 0x23,
    RecurrenceId: 0x24,
    Reminder: 0x25,
    ResponseRequested: 0x26,
    Recurrences: 0x27,
    Recurrence: 0x28,
    Type: 0x29,
    Until: 0x2a,
    Occurrences: 0x2b,
    Interval: 0x2c,
    DayOfWeek: 0x2d,
    DayOfMonth: 0x2e,
    WeekOfMonth: 0x2f,
    MonthOfYear: 0x30,
    StartTime: 0x31,
    Sensitivity: 0x32,
    TimeZone: 0x33,
    GlobalObjId: 0x34,
    ThreadTopic: 0x35,
    MIMEData: 0x36,
    MIMETruncated: 0x37,
    MIMESize: 0x38,
    InternetCPID: 0x39,
    Flag: 0x3a,
    Status: 0x3b,
    ContentClass: 0x3c,
    FlagType: 0x3d,
    CompleteTime: 0x3e,
    DisallowNewTimeProposal: 0x3f,
  }),
  createCodePage("", {}),
  createCodePage("Calendar", {
    TimeZone: 0x05,
    AllDayEvent: 0x06,
    Attendees: 0x07,
    Attendee: 0x08,
    Email: 0x09,
    Name: 0x0a,
    Body: 0x0b,
    BodyTruncated: 0x0c,
    BusyStatus: 0x0d,
    Categories: 0x0e,
    Category: 0x0f,
    Rtf: 0x10,
    DtStamp: 0x11,
    EndTime: 0x12,
    Exception: 0x13,
    Exceptions: 0x14,
    Deleted: 0x15,
    ExceptionStartTime: 0x16,
    Location: 0x17,
    MeetingStatus: 0x18,
    OrganizerEmail: 0x19,
    OrganizerName: 0x1a,
    Recurrence: 0x1b,
    Type: 0x1c,
    Until: 0x1d,
    Occurrences: 0x1e,
    Interval: 0x1f,
    DayOfWeek: 0x20,
    DayOfMonth: 0x21,
    WeekOfMonth: 0x22,
    MonthOfYear: 0x23,
    Reminder: 0x24,
    Sensitivity: 0x25,
    Subject: 0x26,
    StartTime: 0x27,
    UID: 0x28,
    AttendeeStatus: 0x29,
    AttendeeType: 0x2a,
    DisallowNewTimeProposal: 0x33,
    ResponseRequested: 0x34,
    AppointmentReplyTime: 0x35,
    ResponseType: 0x36,
    CalendarType: 0x37,
    IsLeapMonth: 0x38,
    FirstDayOfWeek: 0x39,
    OnlineMeetingConfLink: 0x3a,
    OnlineMeetingExternalLink: 0x3b,
    ClientUid: 0x3c,
  }),
  createCodePage("Move", {}),
  createCodePage("GetItemEstimate", {
    GetItemEstimate: 0x05,
    Version: 0x06,
    Collections: 0x07,
    Collection: 0x08,
    Class: 0x09,
    CollectionId: 0x0a,
    DateTime: 0x0b,
    Estimate: 0x0c,
    Response: 0x0d,
    Status: 0x0e,
  }),
  createCodePage("FolderHierarchy", {
    DisplayName: 0x07,
    ServerId: 0x08,
    ParentId: 0x09,
    Type: 0x0a,
    Status: 0x0c,
    Changes: 0x0e,
    Add: 0x0f,
    Delete: 0x10,
    Update: 0x11,
    SyncKey: 0x12,
    FolderCreate: 0x13,
    FolderDelete: 0x14,
    FolderUpdate: 0x15,
    FolderSync: 0x16,
    Count: 0x17,
  }),
  createCodePage("MeetingResponse", {}),
  createCodePage("Tasks", {}),
  createCodePage("ResolveRecipients", {}),
  createCodePage("ValidateCert", {}),
  createCodePage("Contacts2", {}),
  createCodePage("Ping", {}),
  createCodePage("Provision", {
    Provision: 0x05,
    Policies: 0x06,
    Policy: 0x07,
    PolicyType: 0x08,
    PolicyKey: 0x09,
    Data: 0x0a,
    Status: 0x0b,
    RemoteWipe: 0x0c,
    EASProvisionDoc: 0x0d,
    DevicePasswordEnabled: 0x0e,
    AlphanumericDevicePasswordRequired: 0x0f,
    RequireStorageCardEncryption: 0x10,
    PasswordRecoveryEnabled: 0x11,
    AttachmentsEnabled: 0x13,
    MinDevicePasswordLength: 0x14,
    MaxInactivityTimeDeviceLock: 0x15,
    MaxDevicePasswordFailedAttempts: 0x16,
    MaxAttachmentSize: 0x17,
    AllowSimpleDevicePassword: 0x18,
    DevicePasswordExpiration: 0x19,
    DevicePasswordHistory: 0x1a,
    AllowStorageCard: 0x1b,
    AllowCamera: 0x1c,
    RequireDeviceEncryption: 0x1d,
    AllowUnsignedApplications: 0x1e,
    AllowUnsignedInstallationPackages: 0x1f,
    MinDevicePasswordComplexCharacters: 0x20,
    AllowWiFi: 0x21,
    AllowTextMessaging: 0x22,
    AllowPOPIMAPEmail: 0x23,
    AllowBluetooth: 0x24,
    AllowIrDA: 0x25,
    RequireManualSyncWhenRoaming: 0x26,
    AllowDesktopSync: 0x27,
    MaxCalendarAgeFilter: 0x28,
    AllowHTMLEmail: 0x29,
    MaxEmailAgeFilter: 0x2a,
    MaxEmailBodyTruncationSize: 0x2b,
    MaxEmailHTMLBodyTruncationSize: 0x2c,
    RequireSignedSMIMEMessages: 0x2d,
    RequireEncryptedSMIMEMessages: 0x2e,
    RequireSignedSMIMEAlgorithm: 0x2f,
    RequireEncryptionSMIMEAlgorithm: 0x30,
    AllowSMIMEEncryptionAlgorithmNegotiation: 0x31,
    AllowSMIMESoftCerts: 0x32,
    AllowBrowser: 0x33,
    AllowConsumerEmail: 0x34,
    AllowRemoteDesktop: 0x35,
    AllowInternetSharing: 0x36,
    UnapprovedInROMApplicationList: 0x37,
    ApplicationName: 0x38,
    ApprovedApplicationList: 0x39,
    Hash: 0x3a,
  }),
  createCodePage("Search", {}),
  createCodePage("Gal", {}),
  createCodePage("AirSyncBase", {
    BodyPreference: 0x05,
    Type: 0x06,
    TruncationSize: 0x07,
    AllOrNone: 0x08,
    Body: 0x0a,
    Data: 0x0b,
    EstimatedDataSize: 0x0c,
    Truncated: 0x0d,
    Attachments: 0x0e,
    Attachment: 0x0f,
    DisplayName: 0x10,
    FileReference: 0x11,
    Method: 0x12,
    ContentId: 0x13,
    ContentLocation: 0x14,
    IsInline: 0x15,
    NativeBodyType: 0x16,
    ContentType: 0x17,
    Preview: 0x18,
    BodyPartPreference: 0x19,
    BodyPart: 0x1a,
    Status: 0x1b,
  }),
  createCodePage("Settings", {
    Settings: 0x05,
    Status: 0x06,
    Get: 0x07,
    Set: 0x08,
    DeviceInformation: 0x16,
    Model: 0x17,
    IMEI: 0x18,
    FriendlyName: 0x19,
    OS: 0x1a,
    OSLanguage: 0x1b,
    PhoneNumber: 0x1c,
    UserAgent: 0x20,
    MobileOperator: 0x22,
  }),
  createCodePage("DocumentLibrary", {}),
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
  createCodePage("ComposeMail", {}),
  createCodePage("Email2", {
    UmCallerID: 0x05,
    UmUserNotes: 0x06,
    UmAttDuration: 0x07,
    UmAttOrder: 0x08,
    ConversationId: 0x09,
    ConversationIndex: 0x0a,
    LastVerbExecuted: 0x0b,
    LastVerbExecutionTime: 0x0c,
    ReceivedAsBcc: 0x0d,
    Sender: 0x0e,
    CalendarType: 0x0f,
    IsLeapMonth: 0x10,
    AccountId: 0x11,
    FirstDayOfWeek: 0x12,
    MeetingMessageType: 0x13,
  }),
  createCodePage("Notes", {}),
  createCodePage("RightsManagement", {}),
];

const NAMESPACE_TO_PAGE = new Map(CODE_PAGES.map((page, index) => [page.namespace, index]));

export function encodeWbxml(xml) {
  const nodes = parseXml(xml);
  const bytes = [0x03, 0x01, 0x6a, 0x00];
  const state = { currentPage: 0 };

  for (const node of nodes) {
    appendNode(bytes, node, node.namespace || "AirSync", state);
  }

  return Buffer.from(bytes);
}

export function decodeWbxml(buffer) {
  const bytes = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
  let offset = 0;
  offset += 4;

  const state = { currentPage: 0 };
  const root = parseDocument(bytes, { offset }, state);
  if (!root) {
    return '<?xml version="1.0" encoding="utf-8"?>';
  }

  return `<?xml version="1.0" encoding="utf-8"?>${renderNode(root, true)}`;
}

function appendNode(bytes, node, inheritedNamespace, state) {
  if (node.type === "text") {
    const text = node.value.trim();
    if (!text) {
      return;
    }

    bytes.push(STR_I);
    bytes.push(...Buffer.from(text, "utf8"));
    bytes.push(0x00);
    return;
  }

  const namespace = node.namespace || inheritedNamespace;
  const pageIndex = NAMESPACE_TO_PAGE.get(namespace);
  if (pageIndex === undefined) {
    throw new Error(`Unsupported WBXML namespace: ${namespace}`);
  }

  if (state.currentPage !== pageIndex) {
    bytes.push(SWITCH_PAGE, pageIndex);
    state.currentPage = pageIndex;
  }

  const token = CODE_PAGES[pageIndex].tokens[node.name];
  if (!token) {
    throw new Error(`Unsupported WBXML tag: ${namespace}:${node.name}`);
  }

  const hasContent = node.children.length > 0;
  bytes.push(hasContent ? token | TAG_CONTENT : token);

  if (hasContent) {
    for (const child of node.children) {
      appendNode(bytes, child, namespace, state);
    }

    bytes.push(END);
  }
}

function parseDocument(bytes, cursor, state) {
  const stack = [];
  let root = null;

  while (cursor.offset < bytes.length) {
    const token = bytes[cursor.offset++];

    if (token === SWITCH_PAGE) {
      state.currentPage = bytes[cursor.offset++];
      continue;
    }

    if (token === END) {
      const completed = stack.pop();
      if (!completed) {
        continue;
      }

      if (stack.length > 0) {
        stack[stack.length - 1].children.push(completed);
      } else {
        root = completed;
      }
      continue;
    }

    if (token === STR_I) {
      const value = readInlineString(bytes, cursor);
      if (stack.length > 0) {
        stack[stack.length - 1].children.push({ type: "text", value });
      }
      continue;
    }

    if (token === OPAQUE) {
      const length = readMultiByteInt(bytes, cursor);
      const value = bytes.subarray(cursor.offset, cursor.offset + length).toString("utf8");
      cursor.offset += length;
      if (stack.length > 0) {
        stack[stack.length - 1].children.push({ type: "text", value });
      }
      continue;
    }

    const hasContent = (token & TAG_CONTENT) !== 0;
    const tagToken = token & 0x3f;
    const page = CODE_PAGES[state.currentPage];
    if (!page) {
      throw new Error(`Unknown WBXML code page ${state.currentPage}`);
    }
    const name = page.names[tagToken];
    if (!name) {
      throw new Error(`Unknown WBXML token 0x${tagToken.toString(16)} on page ${state.currentPage}`);
    }

    const element = {
      type: "element",
      name,
      namespace: page.namespace,
      children: [],
    };

    if (hasContent) {
      stack.push(element);
    } else if (stack.length > 0) {
      stack[stack.length - 1].children.push(element);
    } else {
      root = element;
    }
  }

  return root;
}

function renderNode(node, includeXmlns = false) {
  if (node.type === "text") {
    return escapeXml(node.value);
  }

  const xmlns = includeXmlns ? ` xmlns="${node.namespace}:"` : "";
  if (node.children.length === 0) {
    return `<${node.name}${xmlns}/>`;
  }

  const children = node.children.map((child) => renderNode(child, false)).join("");

  return `<${node.name}${xmlns}>${children}</${node.name}>`;
}

function parseXml(xml) {
  const document = xml
    .replace(/<\?xml[\s\S]*?\?>/gi, "")
    .replace(/\r?\n/g, "")
    .trim();
  const stack = [];
  const roots = [];
  let index = 0;

  while (index < document.length) {
    if (document[index] === "<") {
      const endIndex = document.indexOf(">", index);
      const rawTag = document.slice(index + 1, endIndex).trim();

      if (rawTag.startsWith("/")) {
        const closed = stack.pop();
        if (!closed) {
          throw new Error("Malformed XML while encoding WBXML");
        }

        if (stack.length > 0) {
          stack[stack.length - 1].children.push(closed);
        } else {
          roots.push(closed);
        }

        index = endIndex + 1;
        continue;
      }

      const selfClosing = rawTag.endsWith("/");
      const tagContent = selfClosing ? rawTag.slice(0, -1).trim() : rawTag;
      const [rawName, ...attributeParts] = splitTag(tagContent);
      const attributes = attributeParts.join(" ");
      const xmlnsMatch = attributes.match(/xmlns(?::[\w-]+)?="([^"]+):"/i);
      const [prefix, localName] = rawName.includes(":") ? rawName.split(":") : ["", rawName];
      const namespace =
        (xmlnsMatch ? xmlnsMatch[1] : "") ||
        (prefix ? prefixToNamespace(prefix) : stack[stack.length - 1]?.namespace) ||
        "AirSync";

      const element = {
        type: "element",
        name: localName,
        namespace,
        children: [],
      };

      if (selfClosing) {
        if (stack.length > 0) {
          stack[stack.length - 1].children.push(element);
        } else {
          roots.push(element);
        }
      } else {
        stack.push(element);
      }

      index = endIndex + 1;
      continue;
    }

    const nextTag = document.indexOf("<", index);
    const text = document.slice(index, nextTag === -1 ? document.length : nextTag);
    if (text.trim() && stack.length > 0) {
      stack[stack.length - 1].children.push({
        type: "text",
        value: decodeEntities(text.trim()),
      });
    }

    index = nextTag === -1 ? document.length : nextTag;
  }

  while (stack.length > 0) {
    const node = stack.pop();
    if (stack.length > 0) {
      stack[stack.length - 1].children.push(node);
    } else {
      roots.push(node);
    }
  }

  return roots;
}

function splitTag(value) {
  return value.match(/(?:[^\s"]+|"[^"]*")+/g) ?? [];
}

function prefixToNamespace(prefix) {
  if (prefix === "airsyncbase") {
    return "AirSyncBase";
  }

  if (prefix === "airsync") {
    return "AirSync";
  }

  return prefix.charAt(0).toUpperCase() + prefix.slice(1);
}

function createCodePage(namespace, tokens) {
  const names = {};
  for (const [name, token] of Object.entries(tokens)) {
    names[token] = name;
  }

  return { namespace, tokens, names };
}

function readInlineString(bytes, cursor) {
  const start = cursor.offset;
  while (bytes[cursor.offset] !== 0x00) {
    cursor.offset += 1;
  }
  const value = bytes.subarray(start, cursor.offset).toString("utf8");
  cursor.offset += 1;
  return value;
}

function readMultiByteInt(bytes, cursor) {
  let result = 0;
  let current: number;

  do {
    current = bytes[cursor.offset++];
    result = (result << 7) | (current & 0x7f);
  } while ((current & 0x80) !== 0);

  return result;
}

function escapeXml(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function decodeEntities(value) {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}
