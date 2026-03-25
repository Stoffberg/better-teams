import { conversationChatKind, conversationTitle } from "@/lib/chat-format";
import { TeamsApiClient } from "@/services/teams/api-client";
import type { Conversation } from "@/services/teams/types";

const TARGET_NAME = "Siphesihle Thomo";

function normalizeName(value: string): string {
  return value.trim().replace(/\s+/g, " ").toLowerCase();
}

function findMatchingDm(
  conversations: Conversation[],
  selfSkypeId: string | undefined,
  targetName: string,
): Conversation {
  const target = normalizeName(targetName);
  const matches = conversations.filter((c) => {
    if (conversationChatKind(c) !== "dm") return false;
    return normalizeName(conversationTitle(c, selfSkypeId)) === target;
  });
  if (matches.length === 1) return matches[0];
  if (matches.length > 1) {
    throw new Error(`Multiple DM conversations matched "${targetName}".`);
  }
  throw new Error(`No DM conversation matched "${targetName}".`);
}

async function main(): Promise<void> {
  const tenantId = process.env.TEAMS_TENANT_ID;
  const client = new TeamsApiClient(tenantId);
  await client.initialize();

  const account = client.account;
  const allConvos = await client.getAllConversations(100);
  const convo = findMatchingDm(
    allConvos.conversations ?? [],
    account.skypeId,
    TARGET_NAME,
  );

  console.error(`Conversation ID: ${convo.id}`);
  console.error(`Self Skype ID: ${account.skypeId}`);

  const res = await client.getMessages(convo.id, 80, 1);

  // Dump all messages with key fields
  for (const m of res.messages) {
    console.log(
      JSON.stringify(
        {
          id: m.id,
          type: m.type,
          messagetype: m.messagetype,
          contenttype: m.contenttype,
          from: m.from,
          imdisplayname: m.imdisplayname,
          composetime: m.composetime,
          originalarrivaltime: m.originalarrivaltime,
          deleted: m.deleted,
          content: m.content,
          properties_deletetime: m.properties?.deletetime,
          properties_edittime: m.properties?.edittime,
          amsreferences: m.amsreferences,
        },
        null,
        2,
      ),
    );
    console.log("---");
  }
}

main().catch((err) => {
  console.error(err instanceof Error ? err.message : String(err));
  process.exitCode = 1;
});
