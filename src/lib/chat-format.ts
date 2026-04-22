import {
  conversationMemberCount,
  extractTeamsMri,
  messageTimestampValue,
  specialThreadTypeFromConversationId,
} from "@/services/teams/normalize";
import type {
  Conversation,
  ConversationMember,
  Message,
  MessageQuoteReference,
} from "@/services/teams/types";
import { parseHtmlDocument } from "./parse-html";

export type MessageReference = { conversationId: string; messageId: string };

export type MessageAttachment = {
  kind: "image" | "file";
  objectUrl: string;
  openUrl: string;
  thumbnailUrl?: string;
  title: string;
  fileName: string;
  fileSize?: number;
  fileExtension?: string;
};

type ParsedConsumptionHorizon = {
  sequenceId: number;
  timestamp: number;
  messageId: string;
};

type MessageReadStatus = "sending" | "sent" | "delivered" | "read";

export type MessageInlinePart =
  | {
      kind: "text";
      text: string;
      bold?: boolean;
      italic?: boolean;
      strike?: boolean;
      code?: boolean;
    }
  | {
      kind: "link";
      text: string;
      href: string;
      bold?: boolean;
      italic?: boolean;
      strike?: boolean;
    }
  | {
      kind: "mention";
      text: string;
      href?: string;
      messageRef?: MessageReference;
      mentionedMri?: string;
      mentionedDisplayName?: string;
      bold?: boolean;
      italic?: boolean;
      strike?: boolean;
    }
  | {
      kind: "code_block";
      text: string;
      language?: string;
    };

const SKIP_MESSAGE_TYPES = new Set(["Typing", "ClearTyping"]);

const TEAMS_CALL_LOG_STUB =
  /Call Logs for Call\s+[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;

export function textLooksLikeTeamsCallLogStub(text: string): boolean {
  return TEAMS_CALL_LOG_STUB.test(text.trim());
}

function coalesceContent(value: string | undefined | null): string {
  return typeof value === "string" ? value : "";
}

function readableActivityToken(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;
  const readable = trimmed
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/[_./]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
  return readable || null;
}

function normalizedSystemSummary(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const normalized = normalizePreviewText(messagePlainText(value));
  if (!normalized) return null;
  if (/^[[{][\s\S]*[\]}]$/.test(normalized)) return null;
  if (isLikelySystemOrCallBlob(normalized)) return null;
  return normalized;
}

function systemMessageSummary(m: Message): string | null {
  const activity = m.properties?.activity;
  if (activity && typeof activity === "object") {
    const actor =
      nonEmptyString(
        (activity as Record<string, unknown>).sourceUserImDisplayName,
      ) ??
      nonEmptyString(m.senderDisplayName) ??
      nonEmptyString(m.imdisplayname);
    const operation =
      readableActivityToken(
        (activity as Record<string, unknown>).activityOperationType,
      ) ??
      readableActivityToken(
        (activity as Record<string, unknown>).activitySubtype,
      ) ??
      readableActivityToken((activity as Record<string, unknown>).activityType);
    const preview = normalizedSystemSummary(
      (activity as Record<string, unknown>).messagePreview,
    );

    const summary = [actor, operation].filter(Boolean).join(" ").trim();
    if (summary && preview) return `${summary}: ${preview}`;
    if (summary || preview) return summary || preview || null;
  }

  if (!isSystemMessageType(m)) return null;
  return normalizedSystemSummary(m.content);
}

function isSystemMessageType(m: Message): boolean {
  return (
    m.type === "ThreadActivity" ||
    (typeof m.messagetype === "string" &&
      ["Event", "Control", "ThreadActivity"].includes(m.messagetype))
  );
}

type InlineMarks = {
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  code?: boolean;
};

type CachedRichDisplayParts = {
  quote: MessageInlinePart[] | null;
  body: MessageInlinePart[];
  quoteRef: MessageReference | null;
  attachments: MessageAttachment[];
} | null;

const richDisplayPartsCache = new WeakMap<
  Message,
  { signature: string; value: CachedRichDisplayParts }
>();

function textOnlyRichParts(text: string): NonNullable<CachedRichDisplayParts> {
  return {
    quote: null,
    body: [{ kind: "text", text }],
    quoteRef: null,
    attachments: [],
  };
}

function richDisplaySignature(message: Message): string {
  return [
    message.id,
    message.deleted ? "1" : "0",
    message.content ?? "",
    message.contenttype ?? "",
    message.messagetype ?? "",
    typeof message.properties?.files === "string"
      ? message.properties.files
      : "",
    typeof message.properties?.activity === "string"
      ? message.properties.activity
      : JSON.stringify(message.properties?.activity ?? null),
  ].join("\x1f");
}

function compactMarks(marks: InlineMarks): InlineMarks {
  return {
    ...(marks.bold ? { bold: true } : {}),
    ...(marks.italic ? { italic: true } : {}),
    ...(marks.strike ? { strike: true } : {}),
    ...(marks.code ? { code: true } : {}),
  };
}

function sameMarks(left: InlineMarks, right: InlineMarks): boolean {
  return (
    Boolean(left.bold) === Boolean(right.bold) &&
    Boolean(left.italic) === Boolean(right.italic) &&
    Boolean(left.strike) === Boolean(right.strike) &&
    Boolean(left.code) === Boolean(right.code)
  );
}

function appendTextPart(
  parts: MessageInlinePart[],
  text: string,
  marks?: InlineMarks,
) {
  if (!text) return;
  const prev = parts[parts.length - 1];
  if (prev?.kind === "text" && sameMarks(prev, marks ?? {})) {
    prev.text += text;
    return;
  }
  parts.push({ kind: "text", text, ...compactMarks(marks ?? {}) });
}

function normalizedMentionLabel(text: string): string {
  const trimmed = text.replace(/\u00a0/g, " ").trim();
  if (!trimmed) return "";
  return trimmed.startsWith("@") ? trimmed : `@${trimmed}`;
}

function isMentionElement(el: Element): boolean {
  const tag = el.tagName.toUpperCase();
  if (tag === "AT") return true;
  const itemType = (el.getAttribute("itemtype") ?? "").toLowerCase();
  if (itemType.includes("mention")) return true;
  const itemProp = (el.getAttribute("itemprop") ?? "").toLowerCase();
  if (itemProp.includes("mention")) return true;
  const dataMention = el.getAttribute("data-mention-id");
  if (typeof dataMention === "string") return true;
  const className = (el.getAttribute("class") ?? "").toLowerCase();
  return /\bmention\b/.test(className);
}

function sanitizeHref(href: string | null): string | null {
  const trimmed = href?.trim();
  if (!trimmed) return null;
  if (/^(https?:|mailto:|tel:|sip:|msteams:)/i.test(trimmed)) return trimmed;
  return null;
}

function firstDescendantText(el: Element, selector: string): string | null {
  const text = el.querySelector(selector)?.textContent?.trim() ?? "";
  return text || null;
}

function firstDescendantAttribute(
  el: Element,
  selector: string,
  names: string[],
): string | null {
  const node = el.querySelector(selector);
  if (!node) return null;
  return attributeValue(node, names);
}

function parseAttachmentFileSize(value: string | null): number | undefined {
  if (!value) return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : undefined;
}

function isNativeAttachmentMessage(message: Message): boolean {
  const rawMessage = message as unknown as Record<string, unknown>;
  const messageType = (message.messagetype ?? "").trim().toLowerCase();
  const contentType = (message.contenttype ?? "").trim().toLowerCase();
  const formatVariant =
    typeof message.properties?.formatVariant === "string"
      ? message.properties.formatVariant.trim().toLowerCase()
      : "";
  const amsReferences = Array.isArray(rawMessage.amsreferences)
    ? (rawMessage.amsreferences as unknown[])
    : [];

  if (
    messageType === "richtext/media_genericfile" ||
    messageType === "richtext/uriobject" ||
    messageType === "richtext/media_card" ||
    messageType === "richtext/media_callrecording"
  ) {
    return true;
  }
  if (
    contentType === "richtext/media_genericfile" ||
    contentType === "richtext/uriobject"
  ) {
    return true;
  }
  if (
    formatVariant === "richtext/media_genericfile" ||
    formatVariant === "richtext/uriobject"
  ) {
    return true;
  }
  if (amsReferences.length > 0) return true;
  // Fallback: detect <URIObject> in content even when type flags are missing
  const content = coalesceContent(message.content);
  if (/<URIObject[\s>]/i.test(content)) return true;
  return false;
}

function fileExtensionFromName(name: string): string | undefined {
  const dot = name.lastIndexOf(".");
  if (dot < 0 || dot === name.length - 1) return undefined;
  return name.slice(dot + 1).toLowerCase();
}

const IMAGE_EXTENSIONS = new Set([
  "png",
  "jpg",
  "jpeg",
  "gif",
  "webp",
  "avif",
  "bmp",
  "svg",
  "ico",
  "tiff",
  "tif",
  "heic",
  "heif",
]);

function isImageExtension(ext: string | undefined): boolean {
  return ext != null && IMAGE_EXTENSIONS.has(ext.toLowerCase());
}

function parseSingleUriObject(uriObject: Element): MessageAttachment | null {
  const objectUrl = sanitizeHref(uriObject.getAttribute("uri"));
  if (!objectUrl) return null;
  const anchorHref = sanitizeHref(
    uriObject.querySelector("a")?.getAttribute("href") ?? null,
  );
  const thumbnailUrl = sanitizeHref(uriObject.getAttribute("url_thumbnail"));
  const originalName =
    firstDescendantAttribute(uriObject, "OriginalName, originalname", [
      "v",
      "value",
    ]) ??
    firstDescendantAttribute(uriObject, "meta", [
      "originalName",
      "originalname",
    ]);
  const rawTitle = firstDescendantText(uriObject, "Title, title");
  const rawDescription = firstDescendantText(
    uriObject,
    "Description, description",
  );
  const title =
    originalName?.trim() ||
    rawTitle?.replace(/^Title:\s*/i, "").trim() ||
    rawDescription?.replace(/^Description:\s*/i, "").trim() ||
    anchorHref ||
    objectUrl;
  const fileName = originalName?.trim() || title;
  const type = (uriObject.getAttribute("type") ?? "").toLowerCase();
  const ext = fileExtensionFromName(fileName);
  const kind =
    type.startsWith("picture") || isImageExtension(ext) ? "image" : "file";
  return {
    kind,
    objectUrl,
    openUrl: anchorHref ?? objectUrl,
    ...(thumbnailUrl ? { thumbnailUrl } : {}),
    title,
    fileName,
    ...(ext ? { fileExtension: ext } : {}),
    ...(parseAttachmentFileSize(
      firstDescendantAttribute(uriObject, "FileSize, filesize", ["v", "value"]),
    )
      ? {
          fileSize: parseAttachmentFileSize(
            firstDescendantAttribute(uriObject, "FileSize, filesize", [
              "v",
              "value",
            ]),
          ),
        }
      : {}),
  };
}

/**
 * Parse `properties.files` — SharePoint file card attachments.
 * These are used when a file is shared via OneDrive/SharePoint rather than
 * uploaded inline as a URIObject/AMS blob. The message content is typically
 * empty and all metadata lives in `properties.files` (JSON string or array).
 */
function parsePropertiesFileCards(message: Message): MessageAttachment[] {
  const raw = message.properties?.files;
  if (!raw) return [];
  let items: Record<string, unknown>[];
  if (typeof raw === "string") {
    try {
      const parsed = JSON.parse(raw);
      items = Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  } else if (Array.isArray(raw)) {
    items = raw;
  } else {
    return [];
  }
  if (items.length === 0) return [];
  const results: MessageAttachment[] = [];
  for (const item of items) {
    const fileName =
      typeof item.fileName === "string" ? item.fileName.trim() : "";
    const title = typeof item.title === "string" ? item.title.trim() : fileName;
    if (!title && !fileName) continue;
    const objectUrl =
      typeof item.objectUrl === "string" ? item.objectUrl.trim() : "";
    const fileInfo = item.fileInfo as Record<string, unknown> | undefined;
    const shareUrl =
      typeof fileInfo?.shareUrl === "string" ? fileInfo.shareUrl.trim() : "";
    const fileUrl =
      typeof fileInfo?.fileUrl === "string" ? fileInfo.fileUrl.trim() : "";
    const openUrl = shareUrl || fileUrl || objectUrl;
    if (!openUrl) continue;
    const ext = fileExtensionFromName(fileName || title);
    // Extract preview URL for images (AMS-hosted thumbnail/preview)
    const filePreview = item.filePreview as Record<string, unknown> | undefined;
    const previewUrl =
      typeof filePreview?.previewUrl === "string"
        ? filePreview.previewUrl.trim()
        : "";
    results.push({
      kind: isImageExtension(ext) ? "image" : "file",
      objectUrl: objectUrl || openUrl,
      openUrl,
      title: title || fileName,
      fileName: fileName || title,
      ...(ext ? { fileExtension: ext } : {}),
      ...(previewUrl ? { thumbnailUrl: previewUrl } : {}),
    });
  }
  return results;
}

function parseMessageAttachments(message: Message): MessageAttachment[] {
  // First try URIObject-based attachments (AMS uploads)
  if (isNativeAttachmentMessage(message)) {
    const content = coalesceContent(message.content);
    if (content.includes("<")) {
      const doc = parseHtmlDocument(content);
      const uriObjects = doc.body.querySelectorAll("URIObject, uriobject");
      if (uriObjects.length > 0) {
        const results: MessageAttachment[] = [];
        for (const el of uriObjects) {
          const att = parseSingleUriObject(el);
          if (att) results.push(att);
        }
        if (results.length > 0) return results;
      }
    }
  }
  // Fallback: SharePoint/OneDrive file cards in properties.files
  return parsePropertiesFileCards(message);
}

function isPureAttachmentMarkup(content: string): boolean {
  const trimmed = content.trim();
  // Empty content means the attachment is entirely in properties.files
  if (!trimmed) return true;
  // Handle wrapper divs: strip outer <div>...</div> or <p>...</p> wrappers
  const unwrapped = trimmed
    .replace(/^<(?:div|p|span)[^>]*>\s*/i, "")
    .replace(/\s*<\/(?:div|p|span)>\s*$/i, "")
    .trim();
  const check = unwrapped || trimmed;
  return /^<URIObject[\s>]/i.test(check) && /<\/URIObject>\s*$/i.test(check);
}

function parseMessageReferenceFromHref(
  href: string,
  conversationId?: string,
): MessageReference | null {
  const trimmed = href.trim();
  if (!trimmed) return null;
  const messageIdMatch = trimmed.match(/[;?&]messageid=([^&#/]+)/i);
  if (!messageIdMatch?.[1]) return null;
  const messageId = decodeURIComponent(messageIdMatch[1]).trim();
  if (!messageId) return null;
  const conversationIdMatch = trimmed.match(/\/conversations\/([^;/?#]+)/i);
  const resolvedConversationId = decodeURIComponent(
    conversationIdMatch?.[1] ?? conversationId ?? "",
  ).trim();
  if (!resolvedConversationId) return null;
  return { conversationId: resolvedConversationId, messageId };
}

function attributeValue(el: Element, names: string[]): string | null {
  for (const name of names) {
    const value = el.getAttribute(name);
    if (value?.trim()) return value;
  }
  return null;
}

function extractMentionMri(value: unknown): string | null {
  const extracted = extractTeamsMri(value);
  if (extracted) return extracted;
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed.startsWith("8:") ? trimmed : null;
}

function mentionDisplayName(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed || null;
}

function messageMentionLookup(
  message: Message | undefined,
): Map<string, { mri?: string; displayName?: string }> {
  const lookup = new Map<string, { mri?: string; displayName?: string }>();
  const mentions = message?.properties?.mentions;
  if (!Array.isArray(mentions)) return lookup;
  for (const raw of mentions) {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) continue;
    const mention = raw as Record<string, unknown>;
    const ids = [
      mention.id,
      mention.mentionId,
      mention.itemid,
      mention.itemId,
      mention.index,
    ]
      .map((value) => (value == null ? "" : String(value).trim()))
      .filter(Boolean);
    if (ids.length === 0) continue;
    const mri =
      extractMentionMri(mention.mri) ??
      extractMentionMri(mention.mentionedMri) ??
      extractMentionMri(mention.userMri) ??
      extractMentionMri(mention.userId) ??
      extractMentionMri(mention.memberId);
    const displayName =
      mentionDisplayName(mention.displayName) ??
      mentionDisplayName(mention.name) ??
      mentionDisplayName(mention.text) ??
      mentionDisplayName(mention.mentionDisplayName);
    for (const id of ids) {
      lookup.set(id, {
        ...(mri ? { mri } : {}),
        ...(displayName ? { displayName } : {}),
      });
    }
  }
  return lookup;
}

function parseMessageReferenceFromElement(
  el: Element,
  conversationId?: string,
): MessageReference | null {
  const directMessageId = attributeValue(el, [
    "data-message-id",
    "data-messageid",
    "messageid",
    "data-target-message-id",
  ]);
  if (directMessageId && conversationId?.trim()) {
    return {
      conversationId: conversationId.trim(),
      messageId: directMessageId.trim(),
    };
  }

  const directHref = sanitizeHref(
    attributeValue(el, ["href", "data-href", "data-item-url", "itemurl"]),
  );
  if (directHref) {
    const ref = parseMessageReferenceFromHref(directHref, conversationId);
    if (ref) return ref;
  }

  const anchor = el.querySelector("a[href]");
  const nestedHref = sanitizeHref(anchor?.getAttribute("href") ?? null);
  if (!nestedHref) return null;
  return parseMessageReferenceFromHref(nestedHref, conversationId);
}

function messageReferenceFromQuoteReference(
  quoteRef: MessageQuoteReference | undefined,
  conversationId: string,
): MessageReference | null {
  const rawMessageId = quoteRef?.messageId;
  if (rawMessageId == null) return null;
  const messageId = String(rawMessageId).trim();
  if (!messageId) return null;
  return { conversationId, messageId };
}

function messageQuoteReference(m: Message): MessageReference | null {
  const qtdMsgs = m.properties?.qtdMsgs;
  if (Array.isArray(qtdMsgs)) {
    const ref = messageReferenceFromQuoteReference(
      qtdMsgs[0],
      m.conversationId,
    );
    if (ref) return ref;
  }
  return null;
}

function inlinePartsToText(parts: MessageInlinePart[]): string {
  return parts.map((part) => part.text).join("");
}

function mentionIdentity(part: MessageInlinePart): string | null {
  if (part.kind !== "mention") return null;
  if (part.messageRef) {
    return `msg:${part.messageRef.conversationId}:${part.messageRef.messageId}`;
  }
  if (part.mentionedMri) return `mri:${part.mentionedMri}`;
  if (part.href) return `href:${part.href}`;
  return null;
}

function normalizeInlineParts(parts: MessageInlinePart[]): MessageInlinePart[] {
  const normalized: MessageInlinePart[] = [];
  for (const part of parts) {
    if (part.kind === "text") {
      const cleaned = part.text
        .replace(/\u00a0/g, " ")
        .replace(/[ \t]+\n/g, "\n")
        .replace(/\n[ \t]+\n/g, "\n\n")
        .replace(/\n{3,}/g, "\n\n");
      if (!cleaned) continue;
      const prev = normalized[normalized.length - 1];
      if (prev?.kind === "text" && sameMarks(prev, part)) {
        prev.text += cleaned;
        continue;
      }
      normalized.push({ ...compactMarks(part), kind: "text", text: cleaned });
      continue;
    }
    if (!part.text) continue;
    normalized.push(part);
  }
  for (let i = 0; i < normalized.length - 2; i++) {
    const first = normalized[i];
    const middle = normalized[i + 1];
    const third = normalized[i + 2];
    const firstIdentity = first ? mentionIdentity(first) : null;
    const thirdIdentity = third ? mentionIdentity(third) : null;
    if (
      first?.kind === "mention" &&
      middle?.kind === "text" &&
      third?.kind === "mention" &&
      firstIdentity &&
      firstIdentity === thirdIdentity &&
      /^[ \t]+$/.test(middle.text)
    ) {
      normalized.splice(i, 3, {
        ...first,
        text: `${first.text}${middle.text}${third.text.replace(/^@/, "")}`,
        mentionedDisplayName:
          first.mentionedDisplayName ?? third.mentionedDisplayName,
      });
      i -= 1;
    }
  }
  for (let i = 0; i < normalized.length - 1; i++) {
    const current = normalized[i];
    const next = normalized[i + 1];
    if (
      current?.kind === "mention" &&
      current.messageRef &&
      next?.kind === "text"
    ) {
      next.text = next.text.replace(/^\n\n+/, "\n");
      if (!next.text) {
        normalized.splice(i + 1, 1);
        i -= 1;
      }
    }
  }
  return normalized;
}

export function normalizePreviewText(text: string): string {
  let s = text
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  s = s.replace(/\s+Play\s*$/i, "").trim();
  return s;
}

function normalizeThreadBody(text: string): string {
  return text
    .replace(/\u00a0/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\s+Play\s*$/i, "")
    .trim();
}

export function isLikelySystemOrCallBlob(text: string): boolean {
  const s = text.trim();
  if (!s) return true;
  if (textLooksLikeTeamsCallLogStub(s)) return true;
  if (/api\.flightproxy\.teams\.microsoft\.com/i.test(s)) return true;
  if (/\.conv\.skype\.com\/conv\//i.test(s)) return true;
  if (/\{[\s\S]{0,240}"scopeId"[\s\S]{0,240}"callId"/m.test(s)) return true;
  if (/Exception[A-Za-z]{3,}/.test(s)) return true;
  const orgidHits = s.match(/8:orgid:[0-9a-f-]{36}/gi);
  if (orgidHits && orgidHits.length >= 2) return true;
  if (/call(Started|Ended)\b/i.test(s) && /https?:\/\//i.test(s)) return true;
  if (/Recurring\d{2}\/\d{2}\/\d{4}/.test(s)) return true;
  if (/^\s*\{[\s\S]*\}\s*$/.test(s) && s.length > 120) return true;
  if (/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}Recurring/i.test(s)) return true;
  return false;
}

const BLOCK_TAGS = new Set([
  "ADDRESS",
  "ARTICLE",
  "ASIDE",
  "BLOCKQUOTE",
  "BR",
  "DD",
  "DETAILS",
  "DIV",
  "DL",
  "DT",
  "FIELDSET",
  "FIGCAPTION",
  "FIGURE",
  "FOOTER",
  "FORM",
  "H1",
  "H2",
  "H3",
  "H4",
  "H5",
  "H6",
  "HEADER",
  "HR",
  "LI",
  "MAIN",
  "NAV",
  "OL",
  "P",
  "PRE",
  "SECTION",
  "SUMMARY",
  "TABLE",
  "TFOOT",
  "THEAD",
  "TR",
  "UL",
]);

function blockAwareTextContent(el: Element): string {
  const parts: string[] = [];
  function walk(node: Node) {
    if (node.nodeType === 3) {
      parts.push(node.textContent ?? "");
      return;
    }
    if (node.nodeType !== 1) return;
    const tag = (node as Element).tagName;
    if (tag === "BR") {
      parts.push("\n");
      return;
    }
    // Skip URIObject elements — handled separately as attachments
    if (tag === "URIOBJECT") return;
    const isBlock = BLOCK_TAGS.has(tag);
    if (isBlock && parts.length > 0) {
      const last = parts[parts.length - 1];
      if (last && !last.endsWith("\n")) parts.push("\n");
    }
    for (const child of node.childNodes) walk(child);
    if (isBlock) {
      const last = parts[parts.length - 1];
      if (last && !last.endsWith("\n")) parts.push("\n");
    }
  }
  walk(el);
  return parts
    .join("")
    .replace(/\u00a0/g, " ")
    .trim();
}

function blockAwareInlineParts(
  root: ParentNode,
  message?: Message,
  conversationId?: string,
): MessageInlinePart[] {
  const parts: MessageInlinePart[] = [];
  const mentionLookup = messageMentionLookup(message);
  const listStack: Array<{ type: "ul" | "ol"; index: number }> = [];
  function ensureBlockBoundary() {
    const last = parts[parts.length - 1];
    if (!last) return;
    if (last.text.endsWith("\n")) return;
    appendTextPart(parts, "\n");
  }
  function walk(node: Node, marks: InlineMarks = {}) {
    if (node.nodeType === 3) {
      appendTextPart(parts, node.textContent ?? "", marks);
      return;
    }
    if (node.nodeType !== 1) return;
    const el = node as Element;
    const tag = el.tagName.toUpperCase();
    if (tag === "BR") {
      appendTextPart(parts, "\n");
      return;
    }
    // Skip URIObject elements — these are parsed separately as attachments
    if (tag === "URIOBJECT") return;
    const isBlock = BLOCK_TAGS.has(tag);
    if (isBlock) ensureBlockBoundary();
    // Handle <pre> as a code block — extract language and text content
    if (tag === "PRE") {
      const codeEl = el.querySelector("code");
      const codeText = (codeEl ?? el).textContent ?? "";
      // Extract language from class like "language-json" or "language-typescript"
      const langClass =
        el.className?.match?.(/language-(\S+)/)?.[1] ??
        codeEl?.className?.match?.(/language-(\S+)/)?.[1];
      parts.push({
        kind: "code_block",
        text: codeText,
        ...(langClass ? { language: langClass } : {}),
      });
      ensureBlockBoundary();
      return;
    }
    const nextMarks = compactMarks({
      bold: marks.bold || tag === "B" || tag === "STRONG",
      italic: marks.italic || tag === "I" || tag === "EM",
      strike: marks.strike || tag === "S" || tag === "STRIKE" || tag === "DEL",
    });
    // Handle inline <code> (not inside <pre>) with code mark
    if (tag === "CODE") {
      const codeText = el.textContent ?? "";
      if (codeText) {
        appendTextPart(parts, codeText, { ...nextMarks, code: true });
      }
      return;
    }
    if (tag === "UL") {
      listStack.push({ type: "ul", index: 0 });
    }
    if (tag === "OL") {
      listStack.push({ type: "ol", index: 0 });
    }
    if (tag === "LI") {
      const currentList = listStack[listStack.length - 1];
      if (currentList) {
        currentList.index += 1;
        appendTextPart(
          parts,
          currentList.type === "ol" ? `${currentList.index}. ` : "• ",
          nextMarks,
        );
      }
    }
    if (isMentionElement(el)) {
      const text = normalizedMentionLabel(el.textContent ?? "");
      if (text) {
        const mentionMeta = mentionLookup.get(
          attributeValue(el, ["id", "data-mention-id", "itemid", "itemId"]) ??
            "",
        );
        const href = sanitizeHref(
          attributeValue(el, ["href", "data-href", "data-item-url", "itemurl"]),
        );
        const messageRef = parseMessageReferenceFromElement(el, conversationId);
        parts.push({
          kind: "mention",
          text,
          ...nextMarks,
          ...(href ? { href } : {}),
          ...(messageRef ? { messageRef } : {}),
          ...(mentionMeta?.mri ? { mentionedMri: mentionMeta.mri } : {}),
          ...(mentionMeta?.displayName
            ? { mentionedDisplayName: mentionMeta.displayName }
            : {}),
        });
      }
    } else if (tag === "A") {
      const href = sanitizeHref(el.getAttribute("href"));
      const text = (el.textContent ?? "").replace(/\u00a0/g, " ");
      if (href && text) {
        parts.push({ kind: "link", text, href, ...nextMarks });
      } else {
        for (const child of el.childNodes) walk(child, nextMarks);
      }
    } else {
      for (const child of el.childNodes) walk(child, nextMarks);
    }
    if (tag === "UL" || tag === "OL") listStack.pop();
    if (isBlock) ensureBlockBoundary();
  }
  for (const child of root.childNodes) walk(child);
  return normalizeInlineParts(parts);
}

export function messagePlainText(html: string): string {
  if (!html) return "";
  if (!html.includes("<")) return html.trim();
  const doc = parseHtmlDocument(html);
  const text = doc.body.textContent ?? "";
  return text
    .replace(/\u00a0/g, " ")
    .replace(/\s+\n/g, "\n")
    .trim();
}

function parsePlainReplyPrefix(text: string): {
  quote: string | null;
  body: string;
} {
  const raw = text.replace(/\u00a0/g, " ").trim();
  if (!raw) return { quote: null, body: "" };
  const lines = raw.split(/\r?\n/);
  if (lines.length === 2) {
    const a = lines[0].trim();
    const b = lines[1].trim();
    if (
      a.length >= 20 &&
      b.length >= 4 &&
      b.length < 800 &&
      /^(Yes|Yeah|Yea|No|Nope|I |Thanks|Thank you|Sure|Ok|OK|Got it|Will do|Can do|Sounds|Agreed)/i.test(
        b,
      ) &&
      !isLikelySystemOrCallBlob(a) &&
      !isLikelySystemOrCallBlob(b)
    ) {
      return {
        quote: normalizeThreadBody(a),
        body: normalizeThreadBody(b),
      };
    }
  }
  const gtLines: string[] = [];
  const restLines: string[] = [];
  let mode: "gt" | "rest" = "gt";
  for (const line of lines) {
    if (mode === "gt" && /^\s*>/.test(line)) {
      gtLines.push(line.replace(/^\s*>\s?/, ""));
      continue;
    }
    mode = "rest";
    restLines.push(line);
  }
  if (gtLines.length > 0 && restLines.some((l) => l.trim().length > 0)) {
    return {
      quote: normalizeThreadBody(gtLines.join("\n")),
      body: normalizeThreadBody(restLines.join("\n")),
    };
  }
  const paras = raw.split(/\n{3,}/);
  if (paras.length >= 2) {
    const first = paras[0].trim();
    const rest = paras.slice(1).join("\n\n").trim();
    if (
      first.length > 0 &&
      rest.length > 0 &&
      first.length <= 2000 &&
      !isLikelySystemOrCallBlob(first) &&
      !isLikelySystemOrCallBlob(rest)
    ) {
      return {
        quote: normalizeThreadBody(first),
        body: normalizeThreadBody(rest),
      };
    }
  }
  const doubleParas = raw.split(/\n\n+/);
  if (doubleParas.length === 2) {
    const first = doubleParas[0].trim();
    const rest = doubleParas[1].trim();
    if (
      first.length >= 15 &&
      rest.length >= 3 &&
      first.length <= 2000 &&
      !isLikelySystemOrCallBlob(first) &&
      !isLikelySystemOrCallBlob(rest)
    ) {
      return {
        quote: normalizeThreadBody(first),
        body: normalizeThreadBody(rest),
      };
    }
  }
  return { quote: null, body: raw };
}

export function parseMessageQuoteAndBody(html: string | undefined | null): {
  quote: string | null;
  body: string;
} {
  const trimmed = coalesceContent(html).trim();
  if (!trimmed) return { quote: null, body: "" };
  if (!trimmed.includes("<")) {
    return parsePlainReplyPrefix(trimmed);
  }
  const doc = parseHtmlDocument(trimmed);
  const bodyEl = doc.body;
  const candidates = [
    ...bodyEl.querySelectorAll(
      'blockquote, [itemtype*="schema.skype.com/Reply" i], [itemtype*="skype.com/Reply" i]',
    ),
  ];
  const topLevel = candidates.filter(
    (el) => !candidates.some((o) => o !== el && o.contains(el)),
  );
  const quoteChunks: string[] = [];
  for (const el of topLevel) {
    const t = blockAwareTextContent(el);
    if (t) quoteChunks.push(t);
    el.remove();
  }
  const quote =
    quoteChunks.length > 0 ? normalizeThreadBody(quoteChunks.join("\n")) : null;
  const rest = (bodyEl.textContent ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+\n/g, "\n")
    .trim();
  const body = normalizeThreadBody(rest);
  if (!quote && body) {
    const plainSplit = parsePlainReplyPrefix(body);
    if (plainSplit.quote) {
      return { quote: plainSplit.quote, body: plainSplit.body };
    }
  }
  return { quote, body };
}

function parseMessageRichQuoteAndBody(
  html: string | undefined | null,
  message?: Message,
  conversationId?: string,
): {
  quote: MessageInlinePart[] | null;
  body: MessageInlinePart[];
} {
  const trimmed = coalesceContent(html).trim();
  if (!trimmed) return { quote: null, body: [] };
  if (!trimmed.includes("<")) {
    const plain = parsePlainReplyPrefix(trimmed);
    return {
      quote: plain.quote ? [{ kind: "text", text: plain.quote }] : null,
      body: plain.body ? [{ kind: "text", text: plain.body }] : [],
    };
  }
  const doc = parseHtmlDocument(trimmed);
  const bodyEl = doc.body;
  const candidates = [
    ...bodyEl.querySelectorAll(
      'blockquote, [itemtype*="schema.skype.com/Reply" i], [itemtype*="skype.com/Reply" i]',
    ),
  ];
  const topLevel = candidates.filter(
    (el) => !candidates.some((o) => o !== el && o.contains(el)),
  );
  const quoteChunks: MessageInlinePart[][] = [];
  for (const el of topLevel) {
    const parts = blockAwareInlineParts(el, message, conversationId);
    if (inlinePartsToText(parts).trim()) quoteChunks.push(parts);
    el.remove();
  }
  const quote =
    quoteChunks.length > 0
      ? normalizeInlineParts(
          quoteChunks.flatMap((chunk, idx) =>
            idx === 0 ? chunk : [{ kind: "text", text: "\n" }, ...chunk],
          ),
        )
      : null;
  const body = normalizeInlineParts(
    blockAwareInlineParts(bodyEl, message, conversationId),
  );
  if (!quote && inlinePartsToText(body).trim()) {
    const plainSplit = parsePlainReplyPrefix(inlinePartsToText(body));
    if (plainSplit.quote) {
      return {
        quote: [{ kind: "text", text: plainSplit.quote }],
        body: plainSplit.body ? [{ kind: "text", text: plainSplit.body }] : [],
      };
    }
  }
  return { quote, body };
}

function getMessageTextParts(content: string | undefined | null): {
  quote: string | null;
  body: string;
} {
  return parseMessageQuoteAndBody(content);
}

export function isRenderableChatMessage(m: Message): boolean {
  const summary = systemMessageSummary(m);
  if (m.messagetype && SKIP_MESSAGE_TYPES.has(m.messagetype)) return false;
  if (m.deleted) return true;
  if (parseMessageAttachments(m).length > 0) return true;
  if (isSystemMessageType(m)) return Boolean(summary);
  const { quote, body } = getMessageTextParts(m.content);
  const combined = [quote, body].filter(Boolean).join("\n");
  if (!combined.trim()) return Boolean(summary);
  if (isLikelySystemOrCallBlob(combined)) return Boolean(summary);
  return true;
}

export function messagePartsForDisplay(m: Message): {
  quote: string | null;
  body: string;
} | null {
  const summary = systemMessageSummary(m);
  if (!isRenderableChatMessage(m)) return null;
  if (m.deleted) {
    return { quote: null, body: "This message has been deleted." };
  }
  if (isSystemMessageType(m) && summary) {
    return { quote: null, body: summary };
  }
  const attachments = parseMessageAttachments(m);
  if (
    attachments.length > 0 &&
    isPureAttachmentMarkup(coalesceContent(m.content))
  ) {
    return {
      quote: null,
      body: attachments.map((a) => a.title).join(", "),
    };
  }
  const parts = getMessageTextParts(m.content);
  if (!(parts.body ?? "").trim() && !parts.quote && summary) {
    return { quote: null, body: summary };
  }
  if (!(parts.body ?? "").trim() && !parts.quote && attachments.length === 0) {
    return null;
  }
  return parts;
}

export function messageRichPartsForDisplay(m: Message): {
  quote: MessageInlinePart[] | null;
  body: MessageInlinePart[];
  quoteRef: MessageReference | null;
  attachments: MessageAttachment[];
} | null {
  const signature = richDisplaySignature(m);
  const cached = richDisplayPartsCache.get(m);
  if (cached && cached.signature === signature) {
    return cached.value;
  }
  const summary = systemMessageSummary(m);
  if (!isRenderableChatMessage(m)) {
    richDisplayPartsCache.set(m, { signature, value: null });
    return null;
  }
  if (m.deleted) {
    const value = textOnlyRichParts("This message has been deleted.");
    richDisplayPartsCache.set(m, { signature, value });
    return value;
  }
  if (isSystemMessageType(m) && summary) {
    const value = textOnlyRichParts(summary);
    richDisplayPartsCache.set(m, { signature, value });
    return value;
  }
  const attachments = parseMessageAttachments(m);
  if (
    attachments.length > 0 &&
    isPureAttachmentMarkup(coalesceContent(m.content))
  ) {
    const value = {
      quote: null,
      body: [],
      quoteRef: null,
      attachments,
    };
    richDisplayPartsCache.set(m, { signature, value });
    return value;
  }
  const parts = parseMessageRichQuoteAndBody(m.content, m, m.conversationId);
  if (!inlinePartsToText(parts.body).trim() && !parts.quote && summary) {
    const value = {
      ...textOnlyRichParts(summary),
      attachments,
    };
    richDisplayPartsCache.set(m, { signature, value });
    return value;
  }
  if (
    !inlinePartsToText(parts.body).trim() &&
    !parts.quote &&
    attachments.length === 0
  ) {
    richDisplayPartsCache.set(m, { signature, value: null });
    return null;
  }
  const value = {
    ...parts,
    quoteRef: parts.quote ? messageQuoteReference(m) : null,
    attachments,
  };
  richDisplayPartsCache.set(m, { signature, value });
  return value;
}

export function messageBodyForDisplay(m: Message): string | null {
  const parts = messagePartsForDisplay(m);
  if (!parts) return null;
  if (parts.quote && parts.body) {
    return `${parts.quote}\n\n${parts.body}`;
  }
  return parts.body || parts.quote || null;
}

export type ConversationChatKind = "dm" | "group" | "meeting";

function decodedConversationId(conversationId: string): string {
  const trimmed = conversationId.trim();
  if (!trimmed) return "";
  try {
    return decodeURIComponent(trimmed);
  } catch {
    return trimmed;
  }
}

function isUnqConsumerPairConversationId(conversationId: string): boolean {
  const decoded = decodedConversationId(conversationId);
  if (!decoded) return false;
  return /^19:[^@]+@unq\.gbl\.spaces$/i.test(decoded);
}

function looksLikeTeamsThreadScopedConversationId(
  conversationId: string,
): boolean {
  const decoded = decodedConversationId(conversationId);
  if (!decoded) return false;
  return /@thread\.(v2|skype|tacv2)$/i.test(decoded);
}

function rosterCountFromConversation(c: Conversation): number {
  if (typeof c.memberCount === "number" && Number.isFinite(c.memberCount)) {
    return c.memberCount;
  }
  return conversationMemberCount(c);
}

export function conversationChatKind(c: Conversation): ConversationChatKind {
  const tt = (c.threadProperties?.threadType ?? "").toLowerCase();
  const topicRaw = (c.threadProperties?.topic ?? "").trim();
  const topic = topicRaw.toLowerCase();
  const rosterCount = rosterCountFromConversation(c);

  if (
    /\b(meeting|meetup|scheduled|calendar)\b/.test(tt) ||
    /\bteams\s+meeting\b/.test(topic) ||
    /\bmeeting\s+chat\b/.test(topic) ||
    /^meeting\b/i.test(topicRaw)
  ) {
    return "meeting";
  }

  if (isUnqConsumerPairConversationId(c.id)) {
    return "dm";
  }

  if (rosterCount > 2) {
    return "group";
  }

  if (rosterCount === 2) {
    return "dm";
  }

  if (looksLikeTeamsThreadScopedConversationId(c.id)) {
    return "group";
  }

  return "dm";
}

export function conversationKindShortLabel(kind: ConversationChatKind): string {
  if (kind === "dm") return "DM";
  if (kind === "group") return "Group";
  return "Meeting";
}

export function initialsFromLabel(label: string): string {
  const s = label.trim();
  if (!s) return "?";
  const parts = s.split(/\s+/).filter(Boolean);
  if (parts.length === 1) {
    const w = parts[0];
    return w.length >= 2
      ? w.slice(0, 2).toUpperCase()
      : w.slice(0, 1).toUpperCase();
  }
  const a = parts[0][0] ?? "";
  const b = parts[parts.length - 1][0] ?? "";
  return `${a}${b}`.toUpperCase();
}

function nonEmptyString(v: unknown): string | undefined {
  if (typeof v !== "string") return undefined;
  const s = v.trim();
  return s.length > 0 ? s : undefined;
}

function displayNameFromMemberLoose(
  member: ConversationMember,
): string | undefined {
  const m = member as unknown as Record<string, unknown>;
  return (
    nonEmptyString(member.displayName) ??
    nonEmptyString(member.friendlyName) ??
    nonEmptyString(m.display_name) ??
    nonEmptyString(m.shortDisplayName)
  );
}

function dmPeerDisplayNameFromMembers(
  c: Conversation,
  selfSkypeId: string,
): string | undefined {
  if (!Array.isArray(c.members)) return undefined;
  for (const member of c.members) {
    if (!member?.id) continue;
    if (isSelfMessage(member.id, selfSkypeId)) continue;
    const label = displayNameFromMemberLoose(member);
    if (label) {
      return normalizePreviewText(messagePlainText(label)) || undefined;
    }
  }
  return undefined;
}

function groupDisplayNamesFromMembers(
  c: Conversation,
  selfSkypeId?: string,
): string[] {
  if (!Array.isArray(c.members)) return [];
  const seen = new Set<string>();
  const names: string[] = [];
  for (const member of c.members) {
    if (!member) continue;
    if (selfSkypeId && member.id && isSelfMessage(member.id, selfSkypeId)) {
      continue;
    }
    const label = displayNameFromMemberLoose(member);
    if (!label) continue;
    const normalized = normalizePreviewText(messagePlainText(label));
    if (!normalized) continue;
    const key = normalized.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    names.push(normalized);
  }
  return names;
}

function groupTitleFromMembers(
  c: Conversation,
  selfSkypeId?: string,
): string | undefined {
  const names = groupDisplayNamesFromMembers(c, selfSkypeId);
  if (names.length === 0) return undefined;
  const visible = names.slice(0, 3);
  const remaining = names.length - visible.length;
  return remaining > 0
    ? `${visible.join(", ")}, +${remaining}`
    : visible.join(", ");
}

function conversationTitleFromProperties(c: Conversation): string | undefined {
  const keys = [
    "msTeamsThreadName",
    "threadFriendlyName",
    "displayName",
    "friendlyName",
    "activityTitle",
    "title",
    "conversationName",
    "chatName",
  ];
  const read = (o: Record<string, unknown> | undefined): string | undefined => {
    if (!o) return undefined;
    for (const k of keys) {
      const v = o[k];
      if (typeof v === "string") {
        const s = normalizePreviewText(messagePlainText(v.trim()));
        if (s.length > 0) return s;
      }
    }
    return undefined;
  };
  const props = c.properties as Record<string, unknown> | undefined;
  const tp = c.threadProperties as Record<string, unknown> | undefined;
  return read(props) ?? read(tp);
}

export function conversationTitle(
  c: Conversation,
  selfSkypeId?: string,
  peerDisplayNameFromProfile?: string,
): string {
  const rawTopic = c.threadProperties?.topic?.trim();
  if (rawTopic) {
    return normalizePreviewText(messagePlainText(rawTopic)) || "Chat";
  }
  const kind = conversationChatKind(c);
  const fromProps = conversationTitleFromProperties(c);
  if (kind === "dm") {
    const lm = c.lastMessage;
    const name = lm?.senderDisplayName?.trim() || lm?.imdisplayname?.trim();
    if (name && lm) {
      if (!selfSkypeId || !isSelfMessage(lm.from, selfSkypeId)) {
        return name;
      }
    }
    if (selfSkypeId) {
      const peerName = dmPeerDisplayNameFromMembers(c, selfSkypeId);
      if (peerName) return peerName;
    }
    const fromProfile = peerDisplayNameFromProfile?.trim();
    if (fromProfile) {
      const s = normalizePreviewText(messagePlainText(fromProfile));
      if (s) return s;
    }
    if (fromProps) return fromProps;
    return "Direct message";
  }
  if (kind === "group") {
    const fromMembers = groupTitleFromMembers(c, selfSkypeId);
    if (fromMembers) return fromMembers;
    if (fromProps) return fromProps;
    return "Group chat";
  }
  if (kind === "meeting") {
    if (fromProps) return fromProps;
    return "Meeting";
  }
  const fromLm =
    c.lastMessage?.senderDisplayName?.trim() ||
    c.lastMessage?.imdisplayname?.trim();
  if (fromLm) return fromLm;
  return "Chat";
}

export function formatSidebarTime(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  const now = new Date();
  const sod = (x: Date) =>
    new Date(x.getFullYear(), x.getMonth(), x.getDate()).getTime();
  const diffDays = Math.round((sod(now) - sod(d)) / 86_400_000);
  if (diffDays === 0) {
    return new Intl.DateTimeFormat(undefined, {
      hour: "numeric",
      minute: "2-digit",
    }).format(d);
  }
  if (diffDays === 1) return "Yesterday";
  if (diffDays > 1 && diffDays < 7) {
    return new Intl.DateTimeFormat(undefined, { weekday: "short" }).format(d);
  }
  return new Intl.DateTimeFormat(undefined, {
    month: "short",
    day: "numeric",
  }).format(d);
}

export function isSelfMessage(from: string, skypeId?: string): boolean {
  if (!skypeId) return false;
  const tail = skypeId.trim();
  if (!tail) return false;
  const selfMri = tail.startsWith("8:") ? tail : `8:${tail}`;
  const fromMri = extractTeamsMri(from);
  const candidate = (fromMri ?? from.trim()).toLowerCase();
  const selfLower = selfMri.toLowerCase();
  if (candidate === selfLower) return true;
  return tail.length > 0 && candidate.endsWith(tail.toLowerCase());
}

export function formatMessageTime(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  return new Intl.DateTimeFormat(undefined, {
    hour: "numeric",
    minute: "2-digit",
  }).format(d);
}

export function formatDetailedTimestamp(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  return new Intl.DateTimeFormat(undefined, {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(d);
}

export function formatDayLabel(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  const today = new Date();
  const isToday =
    d.getDate() === today.getDate() &&
    d.getMonth() === today.getMonth() &&
    d.getFullYear() === today.getFullYear();
  if (isToday) return "Today";
  return new Intl.DateTimeFormat(undefined, {
    weekday: "long",
    month: "short",
    day: "numeric",
  }).format(d);
}

export function formatThreadDayDividerLabel(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  const today = new Date();
  const isToday =
    d.getDate() === today.getDate() &&
    d.getMonth() === today.getMonth() &&
    d.getFullYear() === today.getFullYear();
  if (isToday) return "Today";
  return new Intl.DateTimeFormat(undefined, {
    weekday: "long",
    month: "long",
    day: "numeric",
  }).format(d);
}

export function isCallLogsStubConversation(c: Conversation): boolean {
  const lm = c.lastMessage;
  if (!lm) return false;
  const { quote, body } = getMessageTextParts(lm.content);
  const combined = [quote, body].filter(Boolean).join("\n");
  return textLooksLikeTeamsCallLogStub(combined);
}

function threadPropertiesAllStringsBlob(
  tp: Record<string, unknown> | undefined,
): string {
  if (!tp) return "";
  const parts: string[] = [];
  for (const v of Object.values(tp)) {
    if (typeof v === "string" && v.trim()) parts.push(v);
  }
  return parts.join("\n");
}

const PIPELINE_EXCLUSION_PROPERTY_KEYS = [
  "msTeamsThreadName",
  "threadFriendlyName",
  "displayName",
  "friendlyName",
  "activityTitle",
  "title",
  "conversationName",
  "chatName",
] as const;

function conversationTitleLikePropertiesBlob(c: Conversation): string {
  const props = c.properties as Record<string, unknown> | undefined;
  if (!props) return "";
  const parts: string[] = [];
  for (const k of PIPELINE_EXCLUSION_PROPERTY_KEYS) {
    const v = props[k];
    if (typeof v === "string" && v.trim()) parts.push(v);
  }
  return parts.join("\n");
}

function pipelineExclusionMetaBlob(c: Conversation): string {
  const tp = c.threadProperties as Record<string, unknown> | undefined;
  return [
    threadPropertiesAllStringsBlob(tp),
    conversationTitleLikePropertiesBlob(c),
  ]
    .filter(Boolean)
    .join("\n");
}

function threadPropertyGroupId(
  tp: Record<string, unknown> | undefined,
): string | undefined {
  if (!tp) return undefined;
  const g = tp.groupId;
  if (typeof g === "string" && g.trim()) return g;
  return undefined;
}

export function isCompanyCommunicationsSidebarThread(c: Conversation): boolean {
  const blob = pipelineExclusionMetaBlob(c).toLowerCase();
  if (!blob) return false;
  if (/\bcompany\s+communications?\b/.test(blob)) return true;
  if (/\bviva\s+company\s+announcements\b/.test(blob)) return true;
  if (/sharepoint\.com\/[^"\s]*companycommunications?/i.test(blob)) return true;
  return false;
}

type AzureAuditPipelineScope = {
  groupIds: Set<string>;
  spaceThreadIds: Set<string>;
};

function isAzureAuditTeamScopeRoot(c: Conversation): boolean {
  const tp = c.threadProperties as Record<string, unknown> | undefined;
  if (!tp) return false;
  const blobLower = threadPropertiesAllStringsBlob(tp).toLowerCase();
  const matchesAudit =
    /azure\s+audit\s+and\s+speciali[sz]ations/.test(blobLower) ||
    /sharepoint\.com\/[^"\s]*azureaudit/i.test(blobLower);
  if (!matchesAudit) return false;
  const ttype = String(tp.threadType ?? "").toLowerCase();
  const product = String(tp.productThreadType ?? "");
  if (ttype === "space") return true;
  if (product === "TeamsTeam" && ttype !== "topic" && ttype !== "meeting") {
    return true;
  }
  return false;
}

function collectAzureAuditPipelineScope(
  conversations: Conversation[],
): AzureAuditPipelineScope {
  const groupIds = new Set<string>();
  const spaceThreadIds = new Set<string>();
  for (const c of conversations) {
    if (!isAzureAuditTeamScopeRoot(c)) continue;
    const tp = c.threadProperties as Record<string, unknown> | undefined;
    const gid = threadPropertyGroupId(tp);
    if (gid) groupIds.add(gid);
    spaceThreadIds.add(c.id);
  }
  return { groupIds, spaceThreadIds };
}

function isAzureExpertMspThreadSignal(blobLower: string): boolean {
  return /\bazure\s+expert\s+msp\b/i.test(blobLower);
}

function isTopicOrSpaceThread(c: Conversation): boolean {
  const tp = c.threadProperties as Record<string, unknown> | undefined;
  if (!tp) return false;
  const threadType = String(tp.threadType ?? "").toLowerCase();
  if (threadType === "topic") return true;
  const spaceId = tp.spaceId;
  return typeof spaceId === "string" && spaceId.trim().length > 0;
}

function lastMessageTextBlob(c: Conversation): string {
  const lm = c.lastMessage;
  if (!lm) return "";
  const { quote, body } = getMessageTextParts(lm.content);
  return messagePlainText([quote, body].filter(Boolean).join("\n"));
}

function isPipelineExcludedConversation(
  c: Conversation,
  azure: AzureAuditPipelineScope,
): boolean {
  if (c.lastMessage == null) return true;
  const specialThreadType =
    c.specialThreadType ?? specialThreadTypeFromConversationId(c.id);
  if (specialThreadType && specialThreadType !== "notes") return true;
  if (isTopicOrSpaceThread(c)) return true;
  if (isCallLogsStubConversation(c)) return true;
  if (isCompanyCommunicationsSidebarThread(c)) return true;
  if (isAzureAuditTeamScopeRoot(c)) return true;
  const blobLower = pipelineExclusionMetaBlob(c).toLowerCase();
  if (isAzureExpertMspThreadSignal(blobLower)) return true;
  const lastLower = lastMessageTextBlob(c).toLowerCase();
  if (isAzureExpertMspThreadSignal(lastLower)) return true;
  const tp = c.threadProperties as Record<string, unknown> | undefined;
  const gid = threadPropertyGroupId(tp);
  if (gid && azure.groupIds.has(gid)) return true;
  const spaceId = tp?.spaceId;
  if (typeof spaceId === "string" && azure.spaceThreadIds.has(spaceId)) {
    return true;
  }
  return false;
}

export function filterConversationsForPipeline(
  conversations: Conversation[],
): Conversation[] {
  const azure = collectAzureAuditPipelineScope(conversations);
  return conversations.filter((c) => !isPipelineExcludedConversation(c, azure));
}

export function includeConversationInSidebar(c: Conversation): boolean {
  return filterConversationsForPipeline([c]).some((x) => x.id === c.id);
}

export function partitionConversationsByKind(list: Conversation[]): {
  meetings: Conversation[];
  groups: Conversation[];
  dms: Conversation[];
} {
  const meetings: Conversation[] = [];
  const groups: Conversation[] = [];
  const dms: Conversation[] = [];
  for (const c of list) {
    const k = conversationChatKind(c);
    if (k === "meeting") meetings.push(c);
    else if (k === "group") groups.push(c);
    else dms.push(c);
  }
  return { meetings, groups, dms };
}

export function messageTimestamp(m: Message): string {
  return messageTimestampValue(m);
}

export function gapBetweenMessages(prev: Message, next: Message): number {
  const ta = Date.parse(messageTimestamp(prev));
  const tb = Date.parse(messageTimestamp(next));
  if (Number.isNaN(ta) || Number.isNaN(tb)) {
    return Number.POSITIVE_INFINITY;
  }
  return tb - ta;
}

export function conversationPreview(c: Conversation): string {
  const lm = c.lastMessage;
  if (!lm) return "No messages yet";
  const summary = systemMessageSummary(lm);
  if (isSystemMessageType(lm) && summary) {
    return summary.length > 72 ? `${summary.slice(0, 72)}…` : summary;
  }
  if (!isRenderableChatMessage(lm)) {
    return "Meeting or system activity";
  }
  const { quote, body } = getMessageTextParts(lm.content);
  if (quote && body) {
    const q = normalizePreviewText(quote);
    const b = normalizePreviewText(body);
    const qShort = q.length > 36 ? `${q.slice(0, 36)}…` : q;
    const merged = `${qShort} · ${b}`;
    return merged.length > 72 ? `${merged.slice(0, 72)}…` : merged;
  }
  const text = normalizePreviewText(quote || body);
  if (!text) return "No messages yet";
  return text.length > 72 ? `${text.slice(0, 72)}…` : text;
}

export function sortConversationsByActivity(
  list: Conversation[],
): Conversation[] {
  return [...list].sort((a, b) => {
    const ta = Date.parse(
      a.lastMessage ? messageTimestampValue(a.lastMessage) : "",
    );
    const tb = Date.parse(
      b.lastMessage ? messageTimestampValue(b.lastMessage) : "",
    );
    const na = Number.isNaN(ta) ? 0 : ta;
    const nb = Number.isNaN(tb) ? 0 : tb;
    return nb - na;
  });
}

export function parseConsumptionHorizon(
  raw: string | undefined | null,
): ParsedConsumptionHorizon | null {
  if (!raw || typeof raw !== "string") return null;
  const parts = raw.split(";");
  if (parts.length < 3) return null;
  const sequenceId = Number(parts[0]);
  const timestamp = Number(parts[1]);
  const messageId = parts.slice(2).join(";").trim();
  if (!Number.isFinite(sequenceId) || !Number.isFinite(timestamp) || !messageId)
    return null;
  return { sequenceId, timestamp, messageId };
}

export function messageReadStatus(
  message: Message,
  peerHorizons: ParsedConsumptionHorizon[],
): MessageReadStatus {
  const seqId =
    message.sequenceId ?? (message.id ? Number(message.id) : Number.NaN);
  if (!Number.isFinite(seqId)) return "sent";
  if (peerHorizons.length === 0) return "sent";
  const anyRead = peerHorizons.some((h) => h.sequenceId >= seqId);
  if (anyRead) return "read";
  return "delivered";
}

export function messageReadTimestamp(
  message: Message,
  peerHorizons: ParsedConsumptionHorizon[],
): string {
  const seqId =
    message.sequenceId ?? (message.id ? Number(message.id) : Number.NaN);
  if (!Number.isFinite(seqId) || peerHorizons.length === 0) return "";
  const matchingHorizon = peerHorizons
    .filter((h) => h.sequenceId >= seqId)
    .sort((left, right) => right.timestamp - left.timestamp)[0];
  if (!matchingHorizon) return "";
  return new Date(matchingHorizon.timestamp).toISOString();
}

/**
 * Check if a message has been edited.
 */
export function isEditedMessage(m: Message): boolean {
  return Boolean(
    m.properties?.edittime &&
      m.properties.edittime !== "0" &&
      m.properties.edittime !== 0,
  );
}
