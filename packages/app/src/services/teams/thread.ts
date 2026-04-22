import type { Message } from "@better-teams/core/teams/types";
import { getTeamsClient } from "./client";

export const teamsThreadService = {
  async getMessages(
    tenantId: string | null | undefined,
    conversationId: string,
    pageSize: number,
    page: number,
  ) {
    const client = await getTeamsClient(tenantId);
    return client.getMessages(conversationId, pageSize, page);
  },

  async getMessagesByUrl(tenantId: string | null | undefined, url: string) {
    const client = await getTeamsClient(tenantId);
    return client.getMessagesByUrl(url);
  },

  async getAnchoredMessages(
    tenantId: string | null | undefined,
    conversationId: string,
    messageId: string,
  ) {
    const client = await getTeamsClient(tenantId);
    return client.getAnchoredMessages(conversationId, messageId);
  },

  async getMembers(
    tenantId: string | null | undefined,
    conversationId: string,
  ) {
    const client = await getTeamsClient(tenantId);
    return client.getThreadMembers(conversationId);
  },

  async getMembersConsumptionHorizon(
    tenantId: string | null | undefined,
    conversationId: string,
  ) {
    const client = await getTeamsClient(tenantId);
    return client.getMembersConsumptionHorizon(conversationId);
  },

  async deleteMessage(
    tenantId: string | null | undefined,
    conversationId: string,
    messageId: string,
  ): Promise<void> {
    const client = await getTeamsClient(tenantId);
    await client.deleteMessage(conversationId, messageId);
  },

  async sendMessage(
    tenantId: string | null | undefined,
    input: {
      conversationId: string;
      content: string;
      contentFormat: "html" | "text";
      mentions: Array<Record<string, unknown>>;
      attachments: File[];
      conversationMembers: string[];
    },
  ): Promise<void> {
    const client = await getTeamsClient(tenantId);
    const displayName = client.account.upn?.split("@")[0]?.trim() || "Me";
    if (input.content.trim()) {
      await client.sendMessage(
        input.conversationId,
        input.content,
        displayName,
        input.contentFormat,
        input.mentions,
      );
    }
    for (const attachment of input.attachments) {
      await client.sendAttachmentMessage(
        input.conversationId,
        attachment,
        displayName,
        input.conversationMembers,
      );
    }
  },

  markDeleted(message: Message): Message {
    return {
      ...message,
      deleted: true,
      content: "",
      properties: {
        ...message.properties,
        deletetime: Date.now(),
      },
    } as Message;
  },
};
