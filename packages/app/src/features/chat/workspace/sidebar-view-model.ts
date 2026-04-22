import {
  conversationChatKind,
  conversationKindShortLabel,
  conversationPreview,
  conversationTitle,
  formatSidebarTime,
  messageTimestamp,
} from "@better-teams/core/chat";
import {
  canonAvatarMri,
  dmConversationAvatarMri,
} from "@better-teams/core/teams/profile/avatars";
import type { Conversation } from "@better-teams/core/teams/types";
import { useMemo } from "react";
import type { SidebarConversationItem } from "../thread/types";

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

export function updateConversationFavoriteState(
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
