import { getAuthToken } from "../src/services/teams/token-extractor";
import type { Conversation } from "../src/services/teams/types";

const AUTHZ_URL = "https://teams.microsoft.com/api/authsvc/v1.0/authz";
const TEAMS_WEB_UA =
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36";

type AuthzJson = {
  tokens?: { skypeToken?: string };
  regionGtms?: { chatService?: string };
};

function threadMetaBlob(c: Conversation): string {
  const tp = c.threadProperties as Record<string, unknown> | undefined;
  if (!tp) return "";
  const parts: string[] = [];
  for (const k of [
    "spaceThreadTopic",
    "topic",
    "topicThreadTopic",
    "sharepointSiteUrl",
    "topics",
  ]) {
    const v = tp[k];
    if (typeof v === "string" && v.trim()) parts.push(v);
  }
  return parts.join("\n");
}

async function main() {
  const needle = process.argv[2];
  if (!needle) {
    console.error("Usage: dump-thread-needle.ts <substring>");
    process.exit(1);
  }
  const rx = new RegExp(needle.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");

  const auth = getAuthToken();
  if (!auth) {
    console.error("No Teams auth token");
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
    console.error("authz failed", authzRes.status);
    process.exit(1);
  }
  const authz = (await authzRes.json()) as AuthzJson;
  const skypeToken = authz.tokens?.skypeToken;
  const chatService = authz.regionGtms?.chatService;
  if (!skypeToken || !chatService) {
    process.exit(1);
  }

  const url = `${chatService}/v1/users/ME/conversations?view=msnp24Equivalent&pageSize=100&startTime=0`;
  const convRes = await fetch(url, {
    headers: { Authentication: `skypetoken=${skypeToken}` },
  });
  if (!convRes.ok) {
    console.error("conversations failed", convRes.status);
    process.exit(1);
  }

  const body = (await convRes.json()) as { conversations?: Conversation[] };
  const list = body.conversations ?? [];

  const hits = list.filter((c) => rx.test(threadMetaBlob(c)));

  console.log("Total conversations:", list.length);
  console.log("Thread-meta hits for", needle, ":", hits.length);
  for (const c of hits) {
    console.log("\n==========");
    console.log("id:", c.id);
    console.log(
      "threadProperties:",
      JSON.stringify(c.threadProperties, null, 2),
    );
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
