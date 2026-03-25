import { conversationChatKind } from "../src/lib/chat-format";
import { getAuthToken } from "../src/services/teams/token-extractor";
import type { Conversation } from "../src/services/teams/types";

const AUTHZ_URL = "https://teams.microsoft.com/api/authsvc/v1.0/authz";
const TEAMS_WEB_UA =
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36";

type AuthzJson = {
  tokens?: { skypeToken?: string };
  regionGtms?: { chatService?: string };
};

async function main() {
  const auth = getAuthToken();
  if (!auth) {
    console.error("No Teams auth token (Teams signed in?)");
    process.exit(1);
  }

  const authzRes = await fetch(AUTHZ_URL, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${auth.token}`,
      "ms-teams-authz-type": "TokenRefresh",
      "User-Agent": TEAMS_WEB_UA,
      Origin: "https://teams.microsoft.com",
      Referer: "https://teams.microsoft.com/",
    },
  });
  if (!authzRes.ok) {
    console.error("authz failed", authzRes.status, await authzRes.text());
    process.exit(1);
  }
  const authz = (await authzRes.json()) as AuthzJson;
  const skypeToken = authz.tokens?.skypeToken;
  const chatService = authz.regionGtms?.chatService;
  if (!skypeToken || !chatService) {
    console.error("authz missing skypeToken or chatService", authz);
    process.exit(1);
  }

  const pageSize = 100;
  const url = `${chatService}/v1/users/ME/conversations?view=msnp24Equivalent&pageSize=${pageSize}&startTime=0`;

  const curlCmd = `curl -sS -H 'Authentication: skypetoken=<skypeToken>' '${url}'`;
  console.log("Equivalent request (token redacted):\n", curlCmd, "\n");

  const convRes = await fetch(url, {
    headers: { Authentication: `skypetoken=${skypeToken}` },
  });
  if (!convRes.ok) {
    console.error("conversations failed", convRes.status, await convRes.text());
    process.exit(1);
  }

  const body = (await convRes.json()) as { conversations?: Conversation[] };
  const list = body.conversations ?? [];

  const needle = (process.argv[2] ?? "engineering").toLowerCase();
  const matches = list.filter((c) => {
    const t = c.threadProperties?.topic ?? "";
    return t.toLowerCase().includes(needle);
  });

  if (matches.length === 0) {
    console.log(
      `No conversation whose threadProperties.topic contains "${needle}".`,
    );
    console.log("Topics in this page (first 40):");
    for (const c of list.slice(0, 40)) {
      const topic = c.threadProperties?.topic ?? "";
      const kind = conversationChatKind(c);
      console.log(
        kind.padEnd(8),
        (c.id ?? "").slice(0, 48),
        topic ? `"${topic.slice(0, 60)}"` : "(no topic)",
      );
    }
    return;
  }

  for (const c of matches) {
    const tp = c.threadProperties;
    const kind = conversationChatKind(c);
    console.log("---");
    console.log("id:", c.id);
    console.log("conversationType:", c.conversationType);
    console.log("inferred kind:", kind);
    console.log("threadProperties:", JSON.stringify(tp ?? {}, null, 2));
    const root = c as Record<string, unknown>;
    const props = root.properties;
    console.log(
      "top-level member-ish:",
      "membercount" in root ? root.membercount : undefined,
      "memberCount" in root ? root.memberCount : undefined,
      "members length:",
      Array.isArray(c.members) ? c.members.length : 0,
    );
    if (props && typeof props === "object") {
      const p = props as Record<string, unknown>;
      console.log(
        "properties member-ish:",
        p.membercount,
        p.memberCount,
        p.participantCount,
      );
    }
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
