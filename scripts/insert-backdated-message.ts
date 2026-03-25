import { conversationChatKind, conversationTitle } from "@/lib/chat-format";
import { TeamsApiClient } from "@/services/teams/api-client";
import type { Conversation } from "@/services/teams/types";

const CLOUD_DIRECT_TENANT = "2d2006bf-2fde-473c-8ce4-ea5307e8eb99";
const TARGET_NAME = "Siphesihle Thomo";
const MESSAGE_TEXT = "Are you ready?";

type ShortProfileLike = {
  displayName?: string | null;
};

type TeamsApiClientInternals = TeamsApiClient & {
  refreshIfNeeded(): Promise<void>;
  regionGtms: { chatService: string } | null;
  skypeToken: string | null;
  httpFetch: (
    input: RequestInfo | URL,
    init?: RequestInit,
  ) => Promise<Response>;
};

function normalizeName(value: string): string {
  return value.trim().replace(/\s+/g, " ").toLowerCase();
}

/**
 * Find a DM conversation by display name.
 *
 * First tries conversations whose title already resolved.  If that fails,
 * falls back to resolving "Direct message" (unresolved) conversations by
 * looking up the other member's profile via fetchShortProfiles.
 */
async function findMatchingDm(
  client: TeamsApiClient,
  conversations: Conversation[],
  selfSkypeId: string | undefined,
  targetName: string,
): Promise<Conversation> {
  const target = normalizeName(targetName);

  // Fast path – title already resolved
  const resolved = conversations.filter((c) => {
    if (conversationChatKind(c) !== "dm") return false;
    return normalizeName(conversationTitle(c, selfSkypeId)) === target;
  });
  if (resolved.length === 1) return resolved[0];
  if (resolved.length > 1) {
    throw new Error(
      `Multiple resolved DMs matched "${targetName}": ${resolved.map((c) => c.id).join(", ")}`,
    );
  }

  // Slow path – resolve unresolved "Direct message" conversations
  const selfGuid = selfSkypeId?.replace(/^.*:/, "") ?? "";
  const dms = conversations.filter(
    (c) =>
      conversationChatKind(c) === "dm" &&
      conversationTitle(c, selfSkypeId) === "Direct message" &&
      c.id.includes("@unq.gbl.spaces"),
  );

  const entries: { conv: Conversation; otherGuid: string }[] = [];
  for (const c of dms) {
    const match = c.id.match(/^19:([a-f0-9-]+)_([a-f0-9-]+)@/);
    if (!match) continue;
    const otherGuid = match[1] === selfGuid ? match[2] : match[1];
    entries.push({ conv: c, otherGuid });
  }

  const mris = entries.map((e) => `8:orgid:${e.otherGuid}`);
  const profiles = await client.fetchShortProfiles(mris);

  for (let i = 0; i < entries.length; i++) {
    const profile = profiles[i] as ShortProfileLike | undefined;
    const displayName = profile?.displayName ?? "";
    if (normalizeName(displayName) === target) {
      return entries[i].conv;
    }
  }

  throw new Error(`No DM conversation matched "${targetName}".`);
}

async function main(): Promise<void> {
  const client = new TeamsApiClient(CLOUD_DIRECT_TENANT);
  await client.initialize();

  const account = client.account;
  process.stderr.write(`Account: ${account.upn}\n`);

  const conversationsResponse = await client.getAllConversations(100);
  const conversation = await findMatchingDm(
    client,
    conversationsResponse.conversations ?? [],
    account.skypeId,
    TARGET_NAME,
  );

  // Build a timestamp exactly one hour ago
  const oneHourAgo = new Date(Date.now() - 60 * 60 * 1000);
  const backdatedIso = oneHourAgo.toISOString();

  // Access the internals to POST directly with a backdated originalarrivaltime
  const internalClient = client as unknown as TeamsApiClientInternals;
  await internalClient.refreshIfNeeded();
  const regionGtms = internalClient.regionGtms;
  const skypeToken = internalClient.skypeToken;
  const httpFetch = internalClient.httpFetch.bind(client);

  const encodedId = encodeURIComponent(conversation.id);
  const url = `${regionGtms.chatService}/v1/users/ME/conversations/${encodedId}/messages`;

  const displayName = account.upn?.trim() || "Better Teams";
  const clientMessageId = oneHourAgo.getTime().toString();

  const body = {
    content: MESSAGE_TEXT,
    messagetype: "Text",
    contenttype: "text",
    amsreferences: [],
    clientmessageid: clientMessageId,
    imdisplayname: displayName,
    originalarrivaltime: backdatedIso,
    composetime: backdatedIso,
    properties: {
      importance: "",
      subject: null,
    },
  };

  const res = await httpFetch(url, {
    method: "POST",
    headers: {
      Authentication: `skypetoken=${skypeToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const errText = await res.text().catch(() => "");
    throw new Error(
      `Send message failed (${res.status}): ${errText.slice(0, 300)}`,
    );
  }

  process.stdout.write(
    `${JSON.stringify(
      {
        sent: true,
        tenantId: account.tenantId ?? null,
        account: account.upn ?? null,
        conversationId: conversation.id,
        targetName: TARGET_NAME,
        message: MESSAGE_TEXT,
        backdatedTo: backdatedIso,
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
