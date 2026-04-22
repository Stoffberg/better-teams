import { getCachedMessages } from "@better-teams/app/services/desktop/runtime";
import { MessagesResponseSchema } from "@better-teams/core/teams/schemas";
import type {
  ConversationMember,
  Message,
} from "@better-teams/core/teams/types";
import { getTeamsClient } from "./client";

export const teamsThreadService = {
  async getMessages(
    tenantId: string | null | undefined,
    conversationId: string,
    pageSize: number,
    page: number,
  ) {
    if (page <= 1) {
      const cached = await getCachedMessages(tenantId, conversationId).catch(
        () => null,
      );
      if (cached && typeof cached === "object") {
        return MessagesResponseSchema.parse(cached);
      }
    }
    try {
      const client = await getTeamsClient(tenantId);
      return client.getMessages(conversationId, pageSize, page);
    } catch (error) {
      if (page > 1) throw error;
      const cached = await getCachedMessages(tenantId, conversationId).catch(
        () => null,
      );
      if (!cached) throw error;
      return MessagesResponseSchema.parse(cached);
    }
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
    const cachedMembers = async () => {
      const cached = await getCachedMessages(tenantId, conversationId).catch(
        () => null,
      );
      if (!cached || typeof cached !== "object") return [];
      return membersFromMessages(MessagesResponseSchema.parse(cached).messages);
    };
    try {
      const client = await getTeamsClient(tenantId);
      const members = filterHumanMembers(
        await client.getThreadMembers(conversationId),
      );
      if (members.length > 0) return members;
      const cached = await cachedMembers();
      return cached.length > 0 ? cached : members;
    } catch (error) {
      const cached = await cachedMembers();
      if (cached.length > 0) return cached;
      throw error;
    }
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

function membersFromMessages(messages: Message[]): ConversationMember[] {
  const seen = new Set<string>();
  return messages.flatMap((message) => {
    if (message.from.trim().toLowerCase().startsWith("28:")) return [];
    const id = (message.fromMri || message.from).trim();
    if (!isHumanMemberId(id)) return [];
    const key = id.toLowerCase();
    if (seen.has(key)) return [];
    seen.add(key);
    const displayName =
      message.senderDisplayName?.trim() || message.imdisplayname?.trim();
    return [
      {
        id,
        role: "User",
        isMri: true,
        ...(displayName ? { displayName } : {}),
      },
    ];
  });
}

function filterHumanMembers(
  members: ConversationMember[],
): ConversationMember[] {
  return members.filter((member) => isHumanMemberId(member.id));
}

function isHumanMemberId(id: string): boolean {
  return id.trim().toLowerCase().startsWith("8:");
}
