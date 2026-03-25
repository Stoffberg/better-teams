/**
 * Debug script: find image attachments in recent messages and test fetching them.
 */

import { execSync } from "node:child_process";
import crypto from "node:crypto";
import { homedir } from "node:os";
import { join } from "node:path";
import Database from "better-sqlite3";

const TENANT_ID = "2d2006bf-2fde-473c-8ce4-ea5307e8eb99";
const CONVERSATION_ID =
  "19:c851229d-64ff-45a9-9228-11e263bea8d5_f4cc62d6-05d5-48b0-9feb-ffe47197d860@unq.gbl.spaces";

function getDecryptionKey() {
  const k = execSync(
    'security find-generic-password -s "Microsoft Teams Safe Storage" -a "Microsoft Teams" -w',
  )
    .toString()
    .trim();
  return crypto.pbkdf2Sync(k, Buffer.from("saltysalt"), 1003, 16, "sha1");
}
function decryptCookie(enc, key) {
  if (!enc?.length) return "";
  if (enc.slice(0, 3).toString() !== "v10") return enc.toString();
  const ct = enc.slice(3);
  if (!ct.length) return "";
  const d = crypto.createDecipheriv("aes-128-cbc", key, Buffer.alloc(16, 0x20));
  return Buffer.concat([d.update(ct), d.final()]).toString("utf8");
}
function extractJwt(raw) {
  const m = raw.match(/(eyJ[A-Za-z0-9_-]+\.eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+)/);
  return m?.[1] ?? "";
}
function decodeJwt(t) {
  try {
    return JSON.parse(Buffer.from(t.split(".")[1], "base64url").toString());
  } catch {
    return null;
  }
}

async function main() {
  const dbPath = join(
    homedir(),
    "Library/Containers/com.microsoft.teams2/Data/Library/Application Support/Microsoft/MSTeams/EBWebView/WV2Profile_tfw/Cookies",
  );
  const key = getDecryptionKey();
  const db = new Database(dbPath, { readonly: true });
  const rows = db
    .prepare(
      `SELECT host_key,name,encrypted_value FROM cookies WHERE (host_key LIKE '%teams%' OR host_key LIKE '%skype%') AND (name='authtoken' OR name='skypetoken_asm') ORDER BY expires_utc DESC`,
    )
    .all();
  db.close();
  const now = Math.floor(Date.now() / 1000);
  let authToken;
  for (const r of rows) {
    const jwt = extractJwt(decryptCookie(r.encrypted_value, key));
    if (!jwt) continue;
    const p = decodeJwt(jwt);
    if (!p || (p.exp ?? 0) < now) continue;
    if (r.name === "authtoken" && p.tid === TENANT_ID) {
      authToken = jwt;
      break;
    }
  }
  if (!authToken) throw new Error("No auth token");

  const c = new AbortController();
  setTimeout(() => c.abort(), 30000);
  const authz = await (
    await fetch("https://teams.microsoft.com/api/authsvc/v1.0/authz", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${authToken}`,
        "Content-Type": "application/json",
      },
      body: "{}",
      signal: c.signal,
    })
  ).json();

  const skypeToken = authz.tokens.skypeToken;
  const chatService = authz.regionGtms.chatService;
  const encoded = encodeURIComponent(CONVERSATION_ID);
  const url = `${chatService}/v1/users/ME/conversations/${encoded}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=80&startTime=1`;

  const c2 = new AbortController();
  setTimeout(() => c2.abort(), 30000);
  const data = await (
    await fetch(url, {
      headers: { Authentication: `skypetoken=${skypeToken}` },
      signal: c2.signal,
    })
  ).json();

  // Find messages with URIObject that look like images
  const imageMessages = [];
  for (const m of data.messages ?? []) {
    const content = m.content || "";
    if (content.includes("URIObject") || content.includes("uriobject")) {
      // Extract type and uri
      const typeMatch = content.match(/type="([^"]+)"/i);
      const uriMatch = content.match(/uri="([^"]+)"/i);
      const thumbMatch = content.match(/url_thumbnail="([^"]+)"/i);
      const nameMatch = content.match(/OriginalName[^>]*v="([^"]+)"/i);
      imageMessages.push({
        id: m.id,
        composetime: m.composetime,
        imdisplayname: m.imdisplayname,
        type: typeMatch?.[1],
        uri: uriMatch?.[1],
        url_thumbnail: thumbMatch?.[1],
        originalName: nameMatch?.[1],
        contentSnippet: content.slice(0, 200),
      });
    }
  }

  console.log(`Found ${imageMessages.length} URIObject messages:\n`);
  for (const msg of imageMessages) {
    console.log(JSON.stringify(msg, null, 2));
    console.log();

    // Try to fetch the image
    if (msg.uri) {
      const viewUrl = `${msg.uri}/views/imgpsh_fullsize_anim`;
      console.log(`  Trying full-size: ${viewUrl}`);
      try {
        const res = await fetch(viewUrl, {
          headers: { Authentication: `skypetoken=${skypeToken}` },
          redirect: "follow",
        });
        console.log(`  Status: ${res.status} ${res.statusText}`);
        console.log(`  Content-Type: ${res.headers.get("content-type")}`);
        console.log(`  Content-Length: ${res.headers.get("content-length")}`);
      } catch (e) {
        console.log(`  Error: ${e.message}`);
      }

      // Also try thumbnail
      if (msg.url_thumbnail) {
        console.log(`  Trying thumbnail: ${msg.url_thumbnail}`);
        try {
          const res = await fetch(msg.url_thumbnail, {
            headers: { Authentication: `skypetoken=${skypeToken}` },
            redirect: "follow",
          });
          console.log(`  Status: ${res.status} ${res.statusText}`);
          console.log(`  Content-Type: ${res.headers.get("content-type")}`);
          console.log(`  Content-Length: ${res.headers.get("content-length")}`);
        } catch (e) {
          console.log(`  Error: ${e.message}`);
        }
      }
    }
    console.log("---");
  }
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
