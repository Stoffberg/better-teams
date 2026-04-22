import { teamsKeys } from "@better-teams/app/lib/teams-query-keys";
import { preloadConversationThread } from "@better-teams/app/lib/teams-thread-preload";
import {
  conversationChatKind,
  conversationKindShortLabel,
  conversationPreview,
  conversationTitle,
  formatSidebarTime,
  messageTimestamp,
} from "@better-teams/core/chat";
import { buildSharedConversationsByMri } from "@better-teams/core/shared-conversation-index";
import type {
  Conversation,
  ConversationMember,
} from "@better-teams/core/teams/types";
import { getOrCreateClient } from "@better-teams/core/teams-client-factory";
import {
  canonAvatarMri,
  dmConversationAvatarMri,
} from "@better-teams/core/teams-profile-avatars";
import { type QueryClient, useMutation, useQuery } from "@tanstack/react-query";
import { useCallback, useEffect, useMemo, useRef } from "react";
import type { SidebarConversationItem } from "./types";

type SidebarViewModelOptions = {
  conversations: Conversation[];
  selfSkypeId?: string;
  avatarThumbByMri: Record<string, string>;
  avatarFullByMri: Record<string, string>;
  displayNameByMri: Record<string, string>;
};

function peerProfileDisplayName(
  conversation: Conversation,
  selfSkypeId: string | undefined,
  byMri: Record<string, string>,
): string | undefined {
  const mri = dmConversationAvatarMri(conversation, selfSkypeId);
  if (!mri) return undefined;
  const displayName = byMri[canonAvatarMri(mri)];
  return typeof displayName === "string" && displayName.trim()
    ? displayName.trim()
    : undefined;
}

function isConversationFavorite(conversation: Conversation): boolean {
  const favorite = conversation.properties?.favorite;
  if (typeof favorite === "boolean") return favorite;
  if (typeof favorite !== "string") return false;
  return favorite.trim().toLowerCase() === "true";
}

function updateConversationFavoriteState(
  conversations: Conversation[],
  conversationId: string,
  favorite: boolean,
): Conversation[] {
  return conversations.map((conversation) =>
    conversation.id === conversationId
      ? {
          ...conversation,
          properties: {
            ...(conversation.properties ?? {}),
            favorite,
          },
        }
      : conversation,
  );
}

export function useSidebarConversationViewModel({
  conversations,
  selfSkypeId,
  avatarThumbByMri,
  avatarFullByMri,
  displayNameByMri,
}: SidebarViewModelOptions): {
  allSidebarItems: SidebarConversationItem[];
  sidebarItemById: Record<string, SidebarConversationItem>;
  sidebarDisplayNameByMri: Record<string, string>;
} {
  const activitySortedConversations = useMemo(
    () =>
      [...conversations].sort((left, right) => {
        const rightTimestamp = Date.parse(
          right.lastMessage ? messageTimestamp(right.lastMessage) : "",
        );
        const leftTimestamp = Date.parse(
          left.lastMessage ? messageTimestamp(left.lastMessage) : "",
        );
        const safeRight = Number.isNaN(rightTimestamp) ? 0 : rightTimestamp;
        const safeLeft = Number.isNaN(leftTimestamp) ? 0 : leftTimestamp;
        if (safeRight !== safeLeft) return safeRight - safeLeft;
        return conversationTitle(left, selfSkypeId).localeCompare(
          conversationTitle(right, selfSkypeId),
        );
      }),
    [conversations, selfSkypeId],
  );

  const allSidebarItems = useMemo<SidebarConversationItem[]>(() => {
    const items = activitySortedConversations.flatMap((conversation) => {
      const title = conversationTitle(
        conversation,
        selfSkypeId,
        peerProfileDisplayName(conversation, selfSkypeId, displayNameByMri),
      );
      const kind = conversationChatKind(conversation);
      if (kind === "dm" && title === "Direct message") return [];
      const preview = conversationPreview(conversation);
      const dmMri = dmConversationAvatarMri(conversation, selfSkypeId);
      const avatarThumbSrc = dmMri
        ? avatarThumbByMri[canonAvatarMri(dmMri)]
        : undefined;
      const avatarFullSrc = dmMri
        ? avatarFullByMri[canonAvatarMri(dmMri)]
        : undefined;
      const lastMessage = conversation.lastMessage;
      const ts = lastMessage ? messageTimestamp(lastMessage) : "";
      return [
        {
          id: conversation.id,
          conversation,
          title,
          preview,
          kind,
          isFavorite: isConversationFavorite(conversation),
          avatarMri: dmMri ? canonAvatarMri(dmMri) : undefined,
          avatarThumbSrc,
          avatarFullSrc,
          sideTime: ts ? formatSidebarTime(ts) : "",
          searchText:
            `${title} ${preview} ${conversationKindShortLabel(kind)}`.toLowerCase(),
        },
      ];
    });

    const favoriteItems = items
      .filter((item) => item.isFavorite)
      .sort((left, right) => left.title.localeCompare(right.title));
    const otherItems = items.filter((item) => !item.isFavorite);

    return [...favoriteItems, ...otherItems];
  }, [
    activitySortedConversations,
    avatarFullByMri,
    avatarThumbByMri,
    displayNameByMri,
    selfSkypeId,
  ]);

  const sidebarItemById = useMemo(
    () => Object.fromEntries(allSidebarItems.map((item) => [item.id, item])),
    [allSidebarItems],
  );
  const sidebarDisplayNameByMri = useMemo(
    () =>
      Object.fromEntries(
        allSidebarItems.flatMap((item) =>
          item.avatarMri ? [[item.avatarMri, item.title] as const] : [],
        ),
      ),
    [allSidebarItems],
  );

  return { allSidebarItems, sidebarItemById, sidebarDisplayNameByMri };
}

export function useFavoriteConversationMutation(
  queryClient: QueryClient,
  activeTenantId?: string | null,
) {
  return useMutation({
    mutationFn: async ({
      conversationId,
      favorite,
    }: {
      conversationId: string;
      favorite: boolean;
    }) => {
      const client = await getOrCreateClient(activeTenantId ?? undefined);
      await client.setConversationFavorite(conversationId, favorite);
      return { conversationId, favorite };
    },
    onMutate: async ({ conversationId, favorite }) => {
      await queryClient.cancelQueries({
        queryKey: teamsKeys.conversations(activeTenantId),
      });
      const previousConversations =
        queryClient.getQueryData<Conversation[]>(
          teamsKeys.conversations(activeTenantId),
        ) ?? [];
      queryClient.setQueryData<Conversation[]>(
        teamsKeys.conversations(activeTenantId),
        updateConversationFavoriteState(
          previousConversations,
          conversationId,
          favorite,
        ),
      );
      return { previousConversations };
    },
    onError: (_error, _variables, context) => {
      if (!context?.previousConversations) return;
      queryClient.setQueryData(
        teamsKeys.conversations(activeTenantId),
        context.previousConversations,
      );
    },
    onSettled: async () => {
      await queryClient.invalidateQueries({
        queryKey: teamsKeys.conversations(activeTenantId),
      });
    },
  });
}

export function useSharedConversationLookup({
  activeTenantId,
  profileSidebarMri,
  allSidebarItems,
  sidebarItemById,
  sidebarDisplayNameByMri,
  displayNameByMri,
  emailByMri,
  queryClient,
}: {
  activeTenantId?: string | null;
  profileSidebarMri: string | null;
  allSidebarItems: SidebarConversationItem[];
  sidebarItemById: Record<string, SidebarConversationItem>;
  sidebarDisplayNameByMri: Record<string, string>;
  displayNameByMri: Record<string, string>;
  emailByMri: Record<string, string>;
  queryClient: QueryClient;
}) {
  const sharedConversationCandidateIds = useMemo(
    () =>
      allSidebarItems
        .filter((item) => item.kind !== "dm")
        .map((item) => item.id)
        .sort(),
    [allSidebarItems],
  );
  const sharedConversationDetailsQuery = useQuery({
    queryKey: [
      "teams",
      "shared-thread-members",
      activeTenantId ?? "__default__",
      profileSidebarMri ?? "__none__",
      sharedConversationCandidateIds.join("\x1f"),
    ],
    queryFn: async () => {
      const byConversationId: Record<string, ConversationMember[]> = {};
      const missingConversationIds: string[] = [];

      for (const conversationId of sharedConversationCandidateIds) {
        const cachedMembers = queryClient.getQueryData<ConversationMember[]>(
          teamsKeys.threadMembers(activeTenantId, conversationId),
        );
        if (cachedMembers && cachedMembers.length > 0) {
          byConversationId[conversationId] = cachedMembers;
          continue;
        }

        const sidebarMembers =
          sidebarItemById[conversationId]?.conversation.members;
        if (sidebarMembers && sidebarMembers.length > 0) {
          byConversationId[conversationId] = sidebarMembers;
          queryClient.setQueryData(
            teamsKeys.threadMembers(activeTenantId, conversationId),
            sidebarMembers,
          );
          continue;
        }

        missingConversationIds.push(conversationId);
      }

      if (missingConversationIds.length === 0) {
        return byConversationId;
      }

      const client = await getOrCreateClient(activeTenantId ?? undefined);
      const concurrency = 6;
      let index = 0;

      await Promise.all(
        Array.from(
          { length: Math.min(concurrency, missingConversationIds.length) },
          async () => {
            while (index < missingConversationIds.length) {
              const conversationId = missingConversationIds[index];
              index += 1;
              if (!conversationId) continue;
              try {
                const members = await client.getThreadMembers(conversationId);
                byConversationId[conversationId] = members;
                queryClient.setQueryData(
                  teamsKeys.threadMembers(activeTenantId, conversationId),
                  members,
                );
              } catch {}
            }
          },
        ),
      );

      return byConversationId;
    },
    enabled:
      Boolean(profileSidebarMri) && sharedConversationCandidateIds.length > 0,
    staleTime: 5 * 60_000,
  });
  const detailedSharedConversationById =
    sharedConversationDetailsQuery.data ?? {};

  return useMemo(
    () =>
      buildSharedConversationsByMri(
        allSidebarItems,
        detailedSharedConversationById,
        { ...sidebarDisplayNameByMri, ...displayNameByMri },
        emailByMri,
      ),
    [
      allSidebarItems,
      detailedSharedConversationById,
      sidebarDisplayNameByMri,
      displayNameByMri,
      emailByMri,
    ],
  );
}

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
