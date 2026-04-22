import { teamsConversationService } from "@better-teams/app/services/teams/conversations";
import { teamsKeys } from "@better-teams/app/services/teams/query-keys";
import type { Conversation } from "@better-teams/core/teams/types";
import { type QueryClient, useMutation } from "@tanstack/react-query";
import { updateConversationFavoriteState } from "./sidebar-view-model";

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
      await teamsConversationService.setFavorite(
        activeTenantId,
        conversationId,
        favorite,
      );
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
