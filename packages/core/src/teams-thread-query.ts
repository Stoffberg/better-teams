import { messageTimestampValue } from "./teams/normalize";
import type { Message, MessagesResponse } from "./teams/types";

export type ThreadQueryData = {
  messages: Message[];
  olderPageUrl: string | null;
  moreOlder: boolean;
};

/**
 * Sort messages by timestamp (ascending / oldest first).
 * Falls back to message id comparison for messages with identical timestamps.
 */
export function sortMessagesByTimestamp(messages: Message[]): Message[] {
  return [...messages].sort((a, b) => {
    const aTs = messageTimestampValue(a);
    const bTs = messageTimestampValue(b);
    if (aTs !== bTs) return aTs.localeCompare(bTs);
    return a.id.localeCompare(b.id);
  });
}

export function threadQueryDataFromResponse(
  res: MessagesResponse,
): ThreadQueryData {
  const list = res.messages ?? [];
  // Always sort by timestamp — the API may return messages out of order
  const sorted = sortMessagesByTimestamp(list);
  const olderPageUrl = res._metadata?.backwardLink ?? null;
  return {
    messages: sorted,
    olderPageUrl,
    moreOlder: olderPageUrl != null && sorted.length > 0,
  };
}
