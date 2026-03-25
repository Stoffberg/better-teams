/**
 * Dump messages from the specific Siphesihle Thomo DM conversation.
 */

import { execSync } from "node:child_process";
import crypto from "node:crypto";
import { homedir } from "node:os";
import { join } from "node:path";
import Database from "better-sqlite3";

const TENANT_ID = "2d2006bf-2fde-473c-8ce4-ea5307e8eb99";
// The DM with Siphesihle Thomo (c851229d = Thomo's guid, f4cc62d6 = Dirk's guid)
const CONVERSATION_ID =
  "19:c851229d-64ff-45a9-9228-11e263bea8d5_f4cc62d6-05d5-48b0-9feb-ffe47197d860@unq.gbl.spaces";

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
  if (encrypted.slice(0, 3).toString() !== "v10") return encrypted.toString();
  const ciphertext = encrypted.slice(3);
  if (ciphertext.length === 0) return "";
  const iv = Buffer.alloc(16, 0x20);
  const decipher = crypto.createDecipheriv("aes-128-cbc", key, iv);
  return Buffer.concat([
    decipher.update(ciphertext),
    decipher.final(),
  ]).toString("utf8");
}

function extractJwt(raw) {
  const jwtPart = "eyJ[A-Za-z0-9_-]+\\.eyJ[A-Za-z0-9_-]+\\.[A-Za-z0-9_-]+";
  for (const prefix of ["Bearer%3D", "Bearer%20", ""]) {
    const m = raw.match(new RegExp(`${prefix}(${jwtPart})`));
    if (m?.[1]) return m[1];
  }
  return "";
}

function decodeJwtPayload(token) {
  try {
    return JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString());
  } catch {
    return null;
  }
}

function extractTokens() {
  const dbPath = join(
    homedir(),
    "Library/Containers/com.microsoft.teams2/Data/Library/Application Support/Microsoft/MSTeams/EBWebView/WV2Profile_tfw/Cookies",
  );
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
    const jwt = extractJwt(decryptCookie(row.encrypted_value, key));
    if (!jwt) continue;
    const payload = decodeJwtPayload(jwt);
    if (!payload || (payload.exp ?? 0) < now) continue;
    tokens.push({ name: row.name, token: jwt, tenantId: payload.tid });
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
  if (!authToken) throw new Error("No auth token");

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

  const encoded = encodeURIComponent(CONVERSATION_ID);
  const url = `${chatService}/v1/users/ME/conversations/${encoded}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=80&startTime=1`;
  const res = await f(url, {
    headers: { Authentication: `skypetoken=${skypeToken}` },
  });
  if (!res.ok) throw new Error(`getMessages failed: ${res.status}`);
  const data = await res.json();
  const messages = data.messages ?? [];

  console.error(`Total messages: ${messages.length}`);
  console.error(
    `Backward link: ${data._metadata?.backwardLink ? "yes" : "no"}`,
  );
  console.error("---");

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
