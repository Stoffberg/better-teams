/**
 * List all DM conversations and their latest messages to find the right one.
 */

import { execSync } from "node:child_process";
import crypto from "node:crypto";
import { existsSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import Database from "better-sqlite3";

const TENANT_ID = "2d2006bf-2fde-473c-8ce4-ea5307e8eb99";
const PBKDF2_SALT = Buffer.from("saltysalt");
const PBKDF2_ITERATIONS = 1003;
const PBKDF2_KEY_LENGTH = 16;

function getDecryptionKey() {
  const safeStorageKey = execSync(
    'security find-generic-password -s "Microsoft Teams Safe Storage" -a "Microsoft Teams" -w',
  )
    .toString()
    .trim();
  return crypto.pbkdf2Sync(
    safeStorageKey,
    PBKDF2_SALT,
    PBKDF2_ITERATIONS,
    PBKDF2_KEY_LENGTH,
    "sha1",
  );
}

function decryptCookie(encrypted, key) {
  if (!encrypted || encrypted.length === 0) return "";
  const prefix = encrypted.slice(0, 3).toString();
  if (prefix !== "v10") return encrypted.toString();
  const ciphertext = encrypted.slice(3);
  if (ciphertext.length === 0) return "";
  const iv = Buffer.alloc(16, 0x20);
  const decipher = crypto.createDecipheriv("aes-128-cbc", key, iv);
  let decrypted = decipher.update(ciphertext);
  decrypted = Buffer.concat([decrypted, decipher.final()]);
  return decrypted.toString("utf8");
}

function extractJwt(raw) {
  const jwtPart = "eyJ[A-Za-z0-9_-]+\\.eyJ[A-Za-z0-9_-]+\\.[A-Za-z0-9_-]+";
  for (const prefix of ["Bearer%3D", "Bearer%20", ""]) {
    const re = new RegExp(`${prefix}(${jwtPart})`);
    const m = raw.match(re);
    if (m?.[1]) return m[1];
  }
  return "";
}

function decodeJwtPayload(token) {
  const parts = token.split(".");
  if (parts.length < 2) return null;
  try {
    return JSON.parse(Buffer.from(parts[1], "base64url").toString());
  } catch {
    return null;
  }
}

function extractTokens() {
  const dbPath = join(
    homedir(),
    "Library/Containers/com.microsoft.teams2/Data/Library/Application Support/Microsoft/MSTeams/EBWebView/WV2Profile_tfw/Cookies",
  );
  if (!existsSync(dbPath)) throw new Error("Teams cookies DB not found");
  const key = getDecryptionKey();
  const db = new Database(dbPath, { readonly: true });
  const rows = db
    .prepare(
      `SELECT host_key, name, encrypted_value FROM cookies WHERE (host_key LIKE '%teams%' OR host_key LIKE '%skype%') AND (name = 'authtoken' OR name = 'skypetoken_asm') ORDER BY expires_utc DESC`,
    )
    .all();
  db.close();
  const now = Math.floor(Date.now() / 1000);
  const tokens = [];
  for (const row of rows) {
    const decrypted = decryptCookie(row.encrypted_value, key);
    const jwt = extractJwt(decrypted);
    if (!jwt) continue;
    const payload = decodeJwtPayload(jwt);
    if (!payload || (payload.exp ?? 0) < now) continue;
    tokens.push({
      name: row.name,
      token: jwt,
      tenantId: payload.tid,
      upn: payload.upn,
      skypeId: payload.skypeid,
    });
  }
  return tokens;
}

async function f(url, opts = {}) {
  const c = new AbortController();
  const t = setTimeout(() => c.abort(), 30000);
  try {
    return await fetch(url, { ...opts, signal: c.signal });
  } finally {
    clearTimeout(t);
  }
}

async function main() {
  const tokens = extractTokens();
  const authToken = tokens.find(
    (t) => t.name === "authtoken" && t.tenantId === TENANT_ID,
  );
  if (!authToken) throw new Error("No auth token for tenant");

  console.error("Account:", authToken.upn, "SkypeId:", authToken.skypeId);

  const authzRes = await f(
    "https://teams.microsoft.com/api/authsvc/v1.0/authz",
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${authToken.token}`,
        "Content-Type": "application/json",
      },
      body: "{}",
    },
  );
  const authz = await authzRes.json();
  const skypeToken = authz.tokens.skypeToken;
  const chatService = authz.regionGtms.chatService;
  const selfSkypeId = authToken.skypeId;

  // Get all conversations
  const convosRes = await f(
    `${chatService}/v1/users/ME/conversations?view=msnp24Equivalent|supportsMessageProperties&pageSize=200&startTime=0&targetType=Passport|Skype|Lync|Thread|NotificationStream|cnsTopicService|Agent`,
    {
      headers: { Authentication: `skypetoken=${skypeToken}` },
    },
  );
  const convos = (await convosRes.json()).conversations ?? [];

  // List all DM conversations with their last message
  const dms = convos.filter((c) => {
    const id = c.id || "";
    return (
      id.includes("@unq.gbl.spaces") ||
      (c.threadProperties?.membercount &&
        Number(c.threadProperties.membercount) <= 2 &&
        !id.includes("@thread"))
    );
  });

  console.log(`\nTotal conversations: ${convos.length}, DMs: ${dms.length}\n`);

  // Sort by last activity (newest first)
  dms.sort((a, b) => {
    const ta = a.lastMessage?.composetime || "";
    const tb = b.lastMessage?.composetime || "";
    return tb.localeCompare(ta);
  });

  // Show each DM with its most recent message
  for (const c of dms.slice(0, 30)) {
    const lm = c.lastMessage;
    const sender = lm?.imdisplayname || "?";
    const time = lm?.composetime || "?";
    const content = (lm?.content || "").slice(0, 100).replace(/\n/g, " ");
    const msgType = lm?.messagetype || "?";
    console.log(`${c.id}`);
    console.log(`  Last: ${time} by ${sender} [${msgType}]`);
    console.log(`  Content: ${content}`);
    console.log();
  }

  // Now try to resolve unresolved DMs by fetching profiles
  const selfGuid = (selfSkypeId || "").replace(/^.*:/, "");
  const unresolvedDms = dms.filter((c) => c.id.includes("@unq.gbl.spaces"));

  console.log(
    `\n--- Resolving ${unresolvedDms.length} DM names via profiles ---\n`,
  );

  // Extract other member GUIDs
  const entries = [];
  for (const c of unresolvedDms) {
    const match = c.id.match(/^19:([a-f0-9-]+)_([a-f0-9-]+)@/);
    if (!match) continue;
    const otherGuid = match[1] === selfGuid ? match[2] : match[1];
    entries.push({ conv: c, otherGuid, mri: `8:orgid:${otherGuid}` });
  }

  // Fetch profiles in batches
  const mris = entries.map((e) => e.mri);
  const profileRes = await f(
    `${authz.regionGtms.middleTier}/beta/users/fetchShortProfile`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${authToken.token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        shortProfileRequests: mris.map((mri) => ({ mri })),
      }),
    },
  );
  const profileData = await profileRes.json();
  const profiles = profileData?.value || profileData || [];

  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];
    const profile = Array.isArray(profiles) ? profiles[i] : null;
    const displayName = profile?.displayName || profile?.givenName || "?";
    const lm = entry.conv.lastMessage;
    console.log(`${entry.conv.id}`);
    console.log(`  Resolved name: ${displayName}`);
    console.log(
      `  Last: ${lm?.composetime || "?"} by ${lm?.imdisplayname || "?"}`,
    );
    console.log(
      `  Content: ${(lm?.content || "").slice(0, 100).replace(/\n/g, " ")}`,
    );

    if (
      displayName.toLowerCase().includes("thomo") ||
      displayName.toLowerCase().includes("siphesihle")
    ) {
      console.log(`  *** MATCH! This is the Siphesihle Thomo DM ***`);
    }
    console.log();
  }
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
