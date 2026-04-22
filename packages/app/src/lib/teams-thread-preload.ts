import { THREAD_PAGE } from "@better-teams/app/components/chat/types";
import { getOrCreateClient } from "@better-teams/core/teams-client-factory";
import {
  type ThreadQueryData,
  threadQueryDataFromResponse,
} from "@better-teams/core/thread";

const preloadInFlight = new Map<string, Promise<ThreadQueryData>>();

function preloadKey(
  tenantId: string | undefined,
  conversationId: string,
): string {
  return `${tenantId ?? "__default__"}\x1f${conversationId}`;
}

export async function preloadConversationThread(
  tenantId: string | undefined,
  conversationId: string,
  _maxCacheAgeMs?: number,
): Promise<ThreadQueryData> {
  const key = preloadKey(tenantId, conversationId);
  const existing = preloadInFlight.get(key);
  if (existing) return existing;

  const request = (async () => {
    const client = await getOrCreateClient(tenantId);
    const response = await client.getMessages(conversationId, THREAD_PAGE, 1);
    return threadQueryDataFromResponse(response);
  })();

  preloadInFlight.set(key, request);
  try {
    return await request;
  } finally {
    preloadInFlight.delete(key);
  }
}
