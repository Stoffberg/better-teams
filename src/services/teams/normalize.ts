import type { Conversation, Message, TeamsSpecialThreadType } from "./types";

export function nonEmptyTrimmedString(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

export function extractTeamsMri(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const input = value.trim();
  if (!input) return null;
  if (input.startsWith("8:")) return input;
  try {
    const decoded = decodeURIComponent(input);
    if (decoded.startsWith("8:")) return decoded;
  } catch {}
  const contactIdx = input.indexOf("/contacts/");
  if (contactIdx >= 0) {
    const tail = input
      .slice(contactIdx + "/contacts/".length)
      .split(/[/?#]/, 1)[0];
    try {
      const decodedTail = decodeURIComponent(tail);
      if (decodedTail.startsWith("8:")) return decodedTail;
    } catch {
      if (tail.startsWith("8:")) return tail;
    }
  }
  const matched = input.match(/8:[^/?#&\s"']+/i);
  return matched ? matched[0] : null;
}

export function normalizeTeamsTimestamp(value: unknown): string {
  if (typeof value === "number" && Number.isFinite(value)) {
    return normalizeTeamsTimestamp(String(value));
  }
  if (typeof value !== "string") return "";
  const trimmed = value.trim();
  if (!trimmed) return "";
  if (/^\d{10,16}$/.test(trimmed)) {
    const raw = Number(trimmed);
    if (Number.isFinite(raw)) {
      const millis = trimmed.length <= 10 ? raw * 1000 : raw;
      const asDate = new Date(millis);
      if (!Number.isNaN(asDate.getTime())) {
        return asDate.toISOString();
      }
    }
  }
  const parsed = new Date(trimmed);
  if (!Number.isNaN(parsed.getTime())) {
    return parsed.toISOString();
  }
  return trimmed;
}

export function parseCountish(value: unknown): number {
  if (value == null) return 0;
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.max(0, Math.floor(value));
  }
  const trimmed = String(value).trim();
  if (!trimmed) return 0;
  const parsed = Number.parseInt(trimmed, 10);
  return Number.isFinite(parsed) ? Math.max(0, parsed) : 0;
}

function parseKnownJsonValue(value: unknown): unknown {
  if (typeof value !== "string") return value;
  const trimmed = value.trim();
  if (!trimmed) return value;
  if (!["{", "["].includes(trimmed[0] ?? "")) return value;
  try {
    return JSON.parse(trimmed) as unknown;
  } catch {
    return value;
  }
}

export function normalizeMessageProperties(
  properties: Record<string, unknown> | undefined,
): Record<string, unknown> | undefined {
  if (!properties) return undefined;
  const next: Record<string, unknown> = { ...properties };
  for (const key of [
    "mentions",
    "files",
    "links",
    "cards",
    "emotions",
    "deltaEmotions",
    "qtdMsgs",
    "pinned",
    "activity",
    "meeting",
    "meta",
    "onbehalfof",
    "atp",
    "botMetadata",
    "botCitations",
    "originalMessageContext",
    "messageUpdatePolicyValue",
    "announceViaEmailPendingMembers",
    "call-log",
  ]) {
    next[key] = parseKnownJsonValue(next[key]);
  }
  return next;
}

export function normalizeConversationProperties(
  properties: Record<string, unknown> | undefined,
): Record<string, unknown> | undefined {
  if (!properties) return undefined;
  const next: Record<string, unknown> = { ...properties };
  for (const key of ["quickReplyAugmentation", "alerts", "meetingInfo"]) {
    next[key] = parseKnownJsonValue(next[key]);
  }
  return next;
}

export function conversationMemberCount(conversation: Conversation): number {
  const thread = conversation.threadProperties;
  const root = conversation as unknown as Record<string, unknown>;
  const props = conversation.properties;
  return Math.max(
    Array.isArray(conversation.members) ? conversation.members.length : 0,
    parseCountish(thread?.membercount),
    parseCountish(thread?.memberCount),
    parseCountish(root.membercount),
    parseCountish(root.memberCount),
    parseCountish(root.participantCount),
    parseCountish(root.participantsCount),
    parseCountish(props?.membercount),
    parseCountish(props?.memberCount),
    parseCountish(props?.participantCount),
    parseCountish(props?.participantsCount),
  );
}

export function messageSenderDisplayName(
  message: Record<string, unknown>,
): string | undefined {
  const direct = nonEmptyTrimmedString(message.imdisplayname);
  if (direct) return direct;
  const tokenDisplay = nonEmptyTrimmedString(message.fromDisplayNameInToken);
  if (tokenDisplay) return tokenDisplay;
  const given = nonEmptyTrimmedString(message.fromGivenNameInToken);
  const family = nonEmptyTrimmedString(message.fromFamilyNameInToken);
  const combined = [given, family].filter(Boolean).join(" ").trim();
  return combined || undefined;
}

export function messageTimestampValue(
  message: Pick<Message, "timestamp" | "originalarrivaltime" | "composetime">,
): string {
  return (
    message.timestamp ||
    message.originalarrivaltime ||
    message.composetime ||
    ""
  );
}

export function conversationLastActivityTime(
  conversation: Pick<Conversation, "lastMessage">,
): string {
  return conversation.lastMessage
    ? messageTimestampValue(conversation.lastMessage)
    : "";
}

export function specialThreadTypeFromConversationId(
  conversationId: string,
): TeamsSpecialThreadType | undefined {
  const trimmed = conversationId.trim().toLowerCase();
  const matched = trimmed.match(
    /^48:(annotations|calllogs|drafts|mentions|notifications|notes)$/,
  );
  return matched?.[1] as TeamsSpecialThreadType | undefined;
}
