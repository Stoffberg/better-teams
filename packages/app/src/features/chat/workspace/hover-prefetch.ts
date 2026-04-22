import { preloadConversationThread } from "@better-teams/app/features/chat/thread/preload";
import { teamsKeys } from "@better-teams/app/services/teams/query-keys";
import type { QueryClient } from "@tanstack/react-query";
import { useCallback, useEffect, useRef } from "react";

export function useConversationHoverPrefetch({
  activeTenantId,
  liveSessionReady,
  activeConversationId,
  queryClient,
}: {
  activeTenantId?: string | null;
  liveSessionReady: boolean;
  activeConversationId: string | null;
  queryClient: QueryClient;
}): {
  handleHoverConversation: (conversationId: string) => void;
  handleHoverConversationEnd: (conversationId: string) => void;
} {
  const hoverPrefetchTimeoutsRef = useRef<Record<string, number>>({});
  const activeConversationIdRef = useRef<string | null>(null);
  activeConversationIdRef.current = activeConversationId;

  useEffect(
    () => () => {
      for (const timeoutId of Object.values(hoverPrefetchTimeoutsRef.current)) {
        window.clearTimeout(timeoutId);
      }
    },
    [],
  );

  const handleHoverConversation = useCallback(
    (conversationId: string) => {
      if (
        !liveSessionReady ||
        conversationId === activeConversationIdRef.current
      ) {
        return;
      }
      if (hoverPrefetchTimeoutsRef.current[conversationId]) return;
      const cachedThreadState = queryClient.getQueryState(
        teamsKeys.thread(activeTenantId, conversationId),
      );
      if (cachedThreadState?.dataUpdatedAt) {
        const ageMs = Date.now() - cachedThreadState.dataUpdatedAt;
        if (ageMs < 60_000) return;
      }
      hoverPrefetchTimeoutsRef.current[conversationId] = window.setTimeout(
        () => {
          delete hoverPrefetchTimeoutsRef.current[conversationId];
          void queryClient
            .prefetchQuery({
              queryKey: teamsKeys.thread(activeTenantId, conversationId),
              queryFn: () =>
                preloadConversationThread(
                  activeTenantId ?? undefined,
                  conversationId,
                  60_000,
                ),
              staleTime: 25_000,
            })
            .catch(() => undefined);
        },
        120,
      );
    },
    [liveSessionReady, activeTenantId, queryClient],
  );
  const handleHoverConversationEnd = useCallback((conversationId: string) => {
    const timeoutId = hoverPrefetchTimeoutsRef.current[conversationId];
    if (!timeoutId) return;
    window.clearTimeout(timeoutId);
    delete hoverPrefetchTimeoutsRef.current[conversationId];
  }, []);

  return { handleHoverConversation, handleHoverConversationEnd };
}
