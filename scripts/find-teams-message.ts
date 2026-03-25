import {
  conversationChatKind,
  conversationKindShortLabel,
  conversationTitle,
  getMessageTextParts,
  messagePlainText,
  messageTimestamp,
} from "@/lib/chat-format";
import { threadQueryDataFromResponse } from "@/lib/teams-thread-query";
import { TeamsApiClient } from "@/services/teams/api-client";
import type { Conversation, Message } from "@/services/teams/types";

function parseArgs(argv: string[]): {
  tenantId: string | undefined;
  needle: string;
  pageSize: number;
  maxMessagePages: number;
} {
  let tenantId: string | undefined;
  let pageSize = 100;
  let maxMessagePages = 80;
  const rest: string[] = [];
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--tenant" && argv[i + 1]) {
      tenantId = argv[i + 1];
      i++;
      continue;
    }
    if (a.startsWith("--tenant=")) {
      tenantId = a.slice("--tenant=".length);
      continue;
    }
    if (a.startsWith("--page-size=")) {
      const n = Number(a.slice("--page-size=".length));
      if (Number.isFinite(n) && n >= 1 && n <= 100) pageSize = n;
      continue;
    }
    if (a.startsWith("--max-message-pages=")) {
      const n = Number(a.slice("--max-message-pages=".length));
      if (Number.isFinite(n) && n >= 1 && n <= 500) maxMessagePages = n;
      continue;
    }
    rest.push(a);
  }
  const needle = rest.join(" ").trim();
  return { tenantId, needle, pageSize, maxMessagePages };
}

function nextConversationStartTime(
  backwardLink: string | undefined,
): number | null {
  if (!backwardLink) return null;
  try {
    const u = new URL(
      backwardLink,
      backwardLink.startsWith("http") ? undefined : "https://placeholder.local",
    );
    const st = u.searchParams.get("startTime");
    if (!st) return null;
    const n = Number(st);
    return Number.isFinite(n) ? n : null;
  } catch {
    const m = backwardLink.match(/[?&]startTime=(\d+)/);
    if (!m) return null;
    const n = Number(m[1]);
    return Number.isFinite(n) ? n : null;
  }
}

async function loadAllConversations(
  client: TeamsApiClient,
  listPageSize: number,
): Promise<Conversation[]> {
  const out: Conversation[] = [];
  let startTime = 0;
  for (;;) {
    const res = await client.getConversationsPageUnfiltered(
      listPageSize,
      startTime,
    );
    const batch = res.conversations ?? [];
    out.push(...batch);
    const next = nextConversationStartTime(res._metadata?.backwardLink);
    if (next == null || batch.length === 0) break;
    startTime = next;
  }
  return out;
}

function messageMatchesNeedle(m: Message, needleLower: string): boolean {
  const { quote, body } = getMessageTextParts(m.content);
  const combined = [quote, body].filter(Boolean).join("\n");
  const plain = messagePlainText(combined);
  return plain.toLowerCase().includes(needleLower);
}

async function findNeedleInConversation(
  client: TeamsApiClient,
  conversationId: string,
  needleLower: string,
  pageSize: number,
  maxMessagePages: number,
): Promise<Message | null> {
  let startTime = 1;
  for (let page = 0; page < maxMessagePages; page++) {
    const res = await client.getMessages(conversationId, pageSize, startTime);
    const td = threadQueryDataFromResponse(res);
    for (const m of td.messages) {
      if (messageMatchesNeedle(m, needleLower)) return m;
    }
    if (!td.moreOlder || td.nextOlderStartTime == null) break;
    startTime = td.nextOlderStartTime;
  }
  return null;
}

async function main(): Promise<void> {
  const envTenant = process.env.TEAMS_TENANT_ID;
  const {
    tenantId: argTenant,
    needle,
    pageSize,
    maxMessagePages,
  } = parseArgs(process.argv.slice(2));
  const tenantId = argTenant ?? envTenant;
  const resolvedNeedle =
    needle ||
    "Azure Expert MSP 2024 for this week's focus sessions, please can you all bring your evidence";
  const needleLower = resolvedNeedle.toLowerCase();

  const client = new TeamsApiClient(tenantId);
  await client.initialize();
  const selfSkypeId = client.account.skypeId;

  process.stderr.write(
    `Account: ${client.account.upn ?? "?"} (tenant ${client.account.tenantId ?? "?"})\n`,
  );
  process.stderr.write(`Searching for: ${JSON.stringify(resolvedNeedle)}\n`);

  const conversations = await loadAllConversations(client, 100);
  process.stderr.write(`Loaded ${conversations.length} conversations.\n`);

  let checked = 0;
  for (const c of conversations) {
    checked++;
    if (checked % 25 === 0) {
      process.stderr.write(`Checked ${checked}/${conversations.length}…\n`);
    }
    const hit = await findNeedleInConversation(
      client,
      c.id,
      needleLower,
      pageSize,
      maxMessagePages,
    );
    if (hit) {
      const kind = conversationChatKind(c);
      const title = conversationTitle(c, selfSkypeId);
      const kindLabel = conversationKindShortLabel(kind);
      const when = messageTimestamp(hit);
      const from = hit.imdisplayname?.trim() || hit.from;
      const parts = getMessageTextParts(hit.content);
      const payload = {
        match: "found",
        conversationId: c.id,
        chatKind: kind,
        chatKindLabel: kindLabel,
        sidebarTitle: title,
        messageId: hit.id,
        messageTime: when,
        from,
        preview: messagePlainText(
          [parts.quote, parts.body].filter(Boolean).join("\n"),
        ).slice(0, 500),
      };
      process.stdout.write(`${JSON.stringify(payload, null, 2)}\n`);
      return;
    }
  }

  process.stdout.write(
    `${JSON.stringify({ match: "not_found", conversationsScanned: conversations.length }, null, 2)}\n`,
  );
  process.exitCode = 1;
}

main().catch((e) => {
  process.stderr.write(`${e instanceof Error ? e.message : String(e)}\n`);
  process.exitCode = 1;
});
