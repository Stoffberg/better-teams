import { conversationChatKind, conversationTitle } from "@/lib/chat-format";
import { TeamsApiClient } from "@/services/teams/api-client";
import type { Conversation } from "@/services/teams/types";

type CliArgs = {
  tenantId?: string;
  targetName: string;
  message?: string;
};

function parseArgs(argv: string[]): CliArgs {
  let tenantId: string | undefined;
  let targetName = "Siphesihle Thomo";
  let message: string | undefined;

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--tenant" && argv[i + 1]) {
      tenantId = argv[i + 1];
      i += 1;
      continue;
    }
    if (arg.startsWith("--tenant=")) {
      tenantId = arg.slice("--tenant=".length);
      continue;
    }
    if (arg === "--to" && argv[i + 1]) {
      targetName = argv[i + 1];
      i += 1;
      continue;
    }
    if (arg.startsWith("--to=")) {
      targetName = arg.slice("--to=".length);
      continue;
    }
    if (arg === "--message" && argv[i + 1]) {
      message = argv[i + 1];
      i += 1;
      continue;
    }
    if (arg.startsWith("--message=")) {
      message = arg.slice("--message=".length);
    }
  }

  return {
    tenantId,
    targetName: targetName.trim(),
    message: message?.trim(),
  };
}

function normalizeName(value: string): string {
  return value.trim().replace(/\s+/g, " ").toLowerCase();
}

function formatLocalTime(date: Date): string {
  return new Intl.DateTimeFormat(undefined, {
    hour: "numeric",
    minute: "2-digit",
  }).format(date);
}

function defaultMessageBody(): string {
  const now = new Date();
  const oneHourAgo = new Date(now.getTime() - 60 * 60 * 1000);
  return `This is a test. Sent at ${formatLocalTime(oneHourAgo)} and backdated by one hour for the demo.`;
}

function findMatchingDm(
  conversations: Conversation[],
  selfSkypeId: string | undefined,
  targetName: string,
): Conversation {
  const target = normalizeName(targetName);
  const matches = conversations.filter((conversation) => {
    if (conversationChatKind(conversation) !== "dm") return false;
    return (
      normalizeName(conversationTitle(conversation, selfSkypeId)) === target
    );
  });

  if (matches.length === 1) {
    return matches[0];
  }

  if (matches.length > 1) {
    const ids = matches.map((conversation) => conversation.id).join(", ");
    throw new Error(
      `Multiple DM conversations matched "${targetName}". Refine the target. Matching conversation ids: ${ids}`,
    );
  }

  throw new Error(`No DM conversation matched "${targetName}".`);
}

async function main(): Promise<void> {
  const envTenant = process.env.TEAMS_TENANT_ID;
  const {
    tenantId: argTenant,
    targetName,
    message: inputMessage,
  } = parseArgs(process.argv.slice(2));
  const tenantId = argTenant ?? envTenant;
  const message = inputMessage || defaultMessageBody();

  if (!targetName) {
    throw new Error("Missing target name.");
  }
  if (!message) {
    throw new Error("Missing message body.");
  }

  const client = new TeamsApiClient(tenantId);
  await client.initialize();

  const account = client.account;
  const conversationsResponse = await client.getAllConversations(100);
  const conversation = findMatchingDm(
    conversationsResponse.conversations ?? [],
    account.skypeId,
    targetName,
  );

  const displayName = account.upn?.trim() || "Better Teams";
  await client.sendMessage(conversation.id, message, displayName);

  process.stdout.write(
    `${JSON.stringify(
      {
        sent: true,
        tenantId: account.tenantId ?? null,
        account: account.upn ?? null,
        conversationId: conversation.id,
        targetName,
        message,
      },
      null,
      2,
    )}\n`,
  );
}

main().catch((error) => {
  process.stderr.write(
    `${error instanceof Error ? error.message : String(error)}\n`,
  );
  process.exitCode = 1;
});
