import { conversationChatKind, conversationTitle } from "@/lib/chat-format";
import { TeamsApiClient } from "@/services/teams/api-client";
import type { Conversation } from "@/services/teams/types";

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
): Promise<Conversation[]> {
  const out: Conversation[] = [];
  let startTime = 0;
  for (;;) {
    const res = await client.getConversationsPageUnfiltered(100, startTime);
    const batch = res.conversations ?? [];
    out.push(...batch);
    const next = nextConversationStartTime(res._metadata?.backwardLink);
    if (next == null || batch.length === 0) break;
    startTime = next;
  }
  return out;
}

function memberSignals(c: Conversation): Record<string, unknown> {
  const root = c as Record<string, unknown>;
  const props =
    root.properties && typeof root.properties === "object"
      ? (root.properties as Record<string, unknown>)
      : undefined;
  const tp =
    c.threadProperties && typeof c.threadProperties === "object"
      ? (c.threadProperties as Record<string, unknown>)
      : undefined;
  return {
    membersLength: Array.isArray(c.members) ? c.members.length : 0,
    top_membercount: root.membercount,
    top_memberCount: root.memberCount,
    top_participantCount: root.participantCount,
    props_membercount: props?.membercount,
    props_memberCount: props?.memberCount,
    props_participantCount: props?.participantCount,
    tp_membercount: tp?.membercount,
    tp_memberCount: tp?.memberCount,
    threadType: tp?.threadType,
    productThreadType: tp?.productThreadType,
    spaceId: tp?.spaceId,
    groupId: tp?.groupId,
    topic: tp?.topic,
    topicThreadTopic: tp?.topicThreadTopic,
    spaceThreadTopic: tp?.spaceThreadTopic,
  };
}

function isLikelyTeamChannel(c: Conversation): boolean {
  const tp =
    c.threadProperties && typeof c.threadProperties === "object"
      ? (c.threadProperties as Record<string, unknown>)
      : undefined;
  const threadType = String(tp?.threadType ?? "").toLowerCase();
  const productThreadType = String(tp?.productThreadType ?? "");
  const hasSpace = typeof tp?.spaceId === "string" && tp.spaceId.length > 0;
  const hasGroup = typeof tp?.groupId === "string" && tp.groupId.length > 0;
  return (
    threadType === "topic" ||
    productThreadType === "TeamsStandardChannel" ||
    hasSpace ||
    hasGroup
  );
}

async function main(): Promise<void> {
  const targetId =
    process.argv[2] ?? "19:f20642e40245468c933870efcf64239a@thread.tacv2";
  const client = new TeamsApiClient(process.env.TEAMS_TENANT_ID);
  await client.initialize();
  const all = await loadAllConversations(client);
  const selfSkypeId = client.account.skypeId;

  const target = all.find((c) => c.id === targetId);
  const firstDm = all.find((c) => conversationChatKind(c) === "dm");
  const firstGroup = all.find((c) => conversationChatKind(c) === "group");

  const payload = {
    account: {
      upn: client.account.upn ?? null,
      tenantId: client.account.tenantId ?? null,
    },
    totals: {
      conversations: all.length,
      dms: all.filter((c) => conversationChatKind(c) === "dm").length,
      groups: all.filter((c) => conversationChatKind(c) === "group").length,
      meetings: all.filter((c) => conversationChatKind(c) === "meeting").length,
    },
    target: target
      ? {
          id: target.id,
          kind: conversationChatKind(target),
          title: conversationTitle(target, selfSkypeId),
          likelyTeamChannel: isLikelyTeamChannel(target),
          memberSignals: memberSignals(target),
        }
      : null,
    normalDmExample: firstDm
      ? {
          id: firstDm.id,
          kind: conversationChatKind(firstDm),
          title: conversationTitle(firstDm, selfSkypeId),
          likelyTeamChannel: isLikelyTeamChannel(firstDm),
          memberSignals: memberSignals(firstDm),
        }
      : null,
    normalGroupExample: firstGroup
      ? {
          id: firstGroup.id,
          kind: conversationChatKind(firstGroup),
          title: conversationTitle(firstGroup, selfSkypeId),
          likelyTeamChannel: isLikelyTeamChannel(firstGroup),
          memberSignals: memberSignals(firstGroup),
        }
      : null,
  };

  process.stdout.write(`${JSON.stringify(payload, null, 2)}\n`);
}

main().catch((e) => {
  process.stderr.write(`${e instanceof Error ? e.message : String(e)}\n`);
  process.exitCode = 1;
});
