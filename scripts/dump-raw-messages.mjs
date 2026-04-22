/**
 * Standalone script to dump raw messages from Teams API.
 * Extracts tokens directly from the cookie store.
 */

import { execSync } from "node:child_process";
import crypto from "node:crypto";
import { existsSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import Database from "better-sqlite3";

// ── Config ──
const TENANT_ID = "2d2006bf-2fde-473c-8ce4-ea5307e8eb99";
const TARGET_NAME = "siphesihle thomo";
const TARGET_TERMS = TARGET_NAME.split(/\s+/)
  .map((term) => term.trim().toLowerCase())
  .filter(Boolean);

// ── Cookie decryption ──
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
  const iv = Buffer.alloc(16, 0x20); // 16 space bytes
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

// ── Extract tokens ──
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
      `SELECT host_key, name, encrypted_value
     FROM cookies
     WHERE (host_key LIKE '%teams%' OR host_key LIKE '%skype%')
       AND (name = 'authtoken' OR name = 'skypetoken_asm')
     ORDER BY expires_utc DESC`,
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
    if (!payload) continue;
    if ((payload.exp ?? 0) < now) continue;
    tokens.push({
      name: row.name,
      host: row.host_key,
      token: jwt,
      tenantId: payload.tid,
      upn: payload.upn,
      skypeId: payload.skypeid,
      audience: payload.aud,
    });
  }
  return tokens;
}

// ── API calls ──
async function callAuthz(bearerToken) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 30000);
  try {
    const res = await fetch(
      "https://teams.microsoft.com/api/authsvc/v1.0/authz",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${bearerToken}`,
          "Content-Type": "application/json",
        },
        body: "{}",
        signal: controller.signal,
      },
    );
    if (!res.ok) throw new Error(`authz failed: ${res.status}`);
    return res.json();
  } finally {
    clearTimeout(timer);
  }
}

async function fetchWithTimeout(url, opts = {}, timeoutMs = 30000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(url, { ...opts, signal: controller.signal });
    return res;
  } finally {
    clearTimeout(timer);
  }
}

async function getMessages(chatService, skypeToken, conversationId) {
  const encoded = encodeURIComponent(conversationId);
  const url = `${chatService}/v1/users/ME/conversations/${encoded}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=80&startTime=1`;
  const res = await fetchWithTimeout(url, {
    headers: { Authentication: `skypetoken=${skypeToken}` },
  });
  if (!res.ok) throw new Error(`getMessages failed: ${res.status}`);
  return res.json();
}

async function getConversations(chatService, skypeToken) {
  const url = `${chatService}/v1/users/ME/conversations?view=msnp24Equivalent|supportsMessageProperties&pageSize=100&startTime=0&targetType=Passport|Skype|Lync|Thread|NotificationStream|cnsTopicService|Agent`;
  const res = await fetchWithTimeout(url, {
    headers: { Authentication: `skypetoken=${skypeToken}` },
  });
  if (!res.ok) throw new Error(`getConversations failed: ${res.status}`);
  return res.json();
}

// ── Main ──
async function main() {
  const tokens = extractTokens();
  const authToken = tokens.find(
    (t) => t.name === "authtoken" && t.tenantId === TENANT_ID,
  );
  if (!authToken) throw new Error("No auth token for tenant");

  console.error("Account:", authToken.upn);

  const authz = await callAuthz(authToken.token);
  const skypeToken = authz.tokens.skypeToken;
  const chatService = authz.regionGtms.chatService;

  console.error("Chat service:", chatService);

  // Find Siphesihle Thomo's conversation
  const convosRes = await getConversations(chatService, skypeToken);
  const convos = convosRes.conversations ?? [];

  // List ALL conversations with siphesihle or thomo in any field
  const candidates = [];
  for (const c of convos) {
    const title = c.threadProperties?.topic || "";
    const members = c.members || [];
    const memberNames = members.map((m) =>
      (m.friendlyName || "").trim().toLowerCase(),
    );
    const lastMsgSender = (c.lastMessage?.imdisplayname || "")
      .trim()
      .toLowerCase();
    const haystack = [title, ...memberNames, lastMsgSender, c.id]
      .join(" ")
      .toLowerCase();

    if (TARGET_TERMS.some((term) => haystack.includes(term))) {
      candidates.push({
        id: c.id,
        lastMsgSender: c.lastMessage?.imdisplayname,
        lastMsgTime: c.lastMessage?.composetime,
        lastMsgSnippet: (c.lastMessage?.content || "").slice(0, 80),
        members: memberNames,
        topic: title,
      });
    }
  }

  console.error("All matching conversations:");
  for (const c of candidates) {
    console.error(JSON.stringify(c, null, 2));
  }

  // Also search for "ooof" or "three-body" or "hello world" in all conversations' last messages
  console.error(
    "\nSearching all conversations for test messages in lastMessage...",
  );
  for (const c of convos) {
    const content = (c.lastMessage?.content || "").toLowerCase();
    if (
      content.includes("ooof") ||
      content.includes("three-body") ||
      content.includes("hello world")
    ) {
      console.error(
        `  MATCH: ${c.id} -> ${c.lastMessage?.imdisplayname}: ${(c.lastMessage?.content || "").slice(0, 100)}`,
      );
    }
  }

  // Pick the first candidate — or if there's one with a target term, prefer it
  let targetConvo = candidates.find((c) =>
    TARGET_TERMS.some((term) => JSON.stringify(c).toLowerCase().includes(term)),
  );
  if (!targetConvo && candidates.length > 0) {
    // Just use the first one
    targetConvo = { id: candidates[0].id };
  }
  if (!targetConvo) {
    throw new Error(`No conversation matched ${TARGET_NAME}`);
  }
  // Convert to the shape we need
  targetConvo = convos.find((c) => c.id === (targetConvo.id ?? targetConvo));

  console.error("Found conversation:", targetConvo.id);

  // Fetch messages
  const msgRes = await getMessages(chatService, skypeToken, targetConvo.id);
  const messages = msgRes.messages ?? [];

  console.error(`Total messages: ${messages.length}`);
  console.error("---");

  // Dump each message
  for (const m of messages) {
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
          content: m.content,
          properties: m.properties
            ? {
                deletetime: m.properties.deletetime,
                edittime: m.properties.edittime,
                hardDeleteTime: m.properties.hardDeleteTime,
              }
            : undefined,
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
  console.error(err);
  process.exitCode = 1;
});
