import { getAuthToken } from "../src/services/teams/token-extractor";
import type { Conversation } from "../src/services/teams/types";

const AUTHZ_URL = "https://teams.microsoft.com/api/authsvc/v1.0/authz";
const TEAMS_WEB_UA =
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36";

type AuthzJson = {
  tokens?: { skypeToken?: string };
  regionGtms?: { chatService?: string };
};

function haystack(c: Conversation): string {
  const root = JSON.stringify(c).toLowerCase();
  return root;
}

async function main() {
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

  const rx =
    /company|communicat|yammer|engage|announcement|org.?wide|broadcast|corporate.?comms/i;
  const hits = list.filter((c) => rx.test(haystack(c)));

  console.log("Total conversations:", list.length);
  console.log(
    "Regex hits (company|communicat|yammer|engage|...):",
    hits.length,
  );
  for (const c of hits) {
    console.log("\n==========");
    console.log("id:", c.id);
    console.log("conversationType:", c.conversationType);
    console.log(
      "threadProperties:",
      JSON.stringify(c.threadProperties, null, 2),
    );
    const p = (c as { properties?: unknown }).properties;
    if (p && typeof p === "object") {
      console.log("properties:", JSON.stringify(p, null, 2));
    }
    const lm = c.lastMessage;
    if (lm) {
      const plain = lm.content
        ?.replace(/<[^>]+>/g, " ")
        .replace(/\s+/g, " ")
        .trim();
      console.log("lastMessage.imdisplayname:", lm.imdisplayname);
      console.log("lastMessage.content snippet:", plain?.slice(0, 120));
    }
  }

  const topicCompany = list.filter((c) =>
    /company\s+communication/i.test(c.threadProperties?.topic ?? ""),
  );
  console.log(
    "\n--- Exact topic match Company communication:",
    topicCompany.length,
  );
  for (const c of topicCompany) {
    console.log(c.id, JSON.stringify(c.threadProperties));
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
