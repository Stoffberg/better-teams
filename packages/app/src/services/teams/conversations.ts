import type { ConversationMember } from "@better-teams/core/teams/types";
import { getTeamsClient } from "./client";

export const teamsConversationService = {
  async list(tenantId?: string | null, limit = 100) {
    const client = await getTeamsClient(tenantId);
    const response = await client.getAllConversations(limit);
    return response.conversations ?? [];
  },

  async setFavorite(
    tenantId: string | null | undefined,
    conversationId: string,
    favorite: boolean,
  ): Promise<void> {
    const client = await getTeamsClient(tenantId);
    await client.setConversationFavorite(conversationId, favorite);
  },

  async getMembers(
    tenantId: string | null | undefined,
    conversationId: string,
  ): Promise<ConversationMember[]> {
    const client = await getTeamsClient(tenantId);
    return client.getThreadMembers(conversationId);
  },
};
