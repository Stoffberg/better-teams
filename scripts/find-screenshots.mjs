/**
 * Find "Screenshot" messages by paging through ALL conversations.
 */

import { execSync } from "node:child_process";
import crypto from "node:crypto";
import { homedir } from "node:os";
import { join } from "node:path";
import Database from "better-sqlite3";

const TENANT_ID = "2d2006bf-2fde-473c-8ce4-ea5307e8eb99";

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

  // Get conversations
  const convUrl = `${chatService}/v1/users/ME/conversations?view=msnp24Equivalent|supportsMessageProperties&pageSize=100&startTime=0&targetType=Passport|Skype|Lync|Thread|NotificationStream|cnsTopicService|Agent`;
  const convData = await (
    await fetch(convUrl, {
      headers: { Authentication: `skypetoken=${skypeToken}` },
    })
  ).json();

  const convos = convData.conversations ?? [];
  console.log(`Total conversations: ${convos.length}\n`);

  for (const conv of convos) {
    const encoded = encodeURIComponent(conv.id);
    const msgUrl = `${chatService}/v1/users/ME/conversations/${encoded}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=80&startTime=1`;
    try {
      const msgData = await (
        await fetch(msgUrl, {
          headers: { Authentication: `skypetoken=${skypeToken}` },
        })
      ).json();

      for (const m of msgData.messages ?? []) {
        const mc = m.content || "";
        const props = m.properties?.files || "";
        const combined = mc + props;
        if (
          combined.toLowerCase().includes("screenshot") ||
          mc.includes('type="Picture')
        ) {
          console.log(`FOUND: ${conv.id}`);
          console.log(
            `  Message ${m.id} from ${m.imdisplayname} at ${m.composetime}`,
          );
          console.log(`  messagetype: ${m.messagetype}`);

          // Extract details
          const typeMatch = mc.match(/type="([^"]+)"/i);
          const uriMatch = mc.match(/uri="([^"]+)"/i);
          const thumbMatch = mc.match(/url_thumbnail="([^"]+)"/i);
          const nameMatch = mc.match(/OriginalName[^>]*v="([^"]+)"/i);
          if (typeMatch) console.log(`  URIObject type: ${typeMatch[1]}`);
          if (uriMatch) console.log(`  URI: ${uriMatch[1]}`);
          if (thumbMatch) console.log(`  Thumbnail: ${thumbMatch[1]}`);
          if (nameMatch) console.log(`  Name: ${nameMatch[1]}`);
          if (props)
            console.log(
              `  properties.files: ${typeof props === "string" ? props.slice(0, 200) : JSON.stringify(props).slice(0, 200)}`,
            );
          console.log(`  Content snippet: ${mc.slice(0, 300)}`);
          console.log();
        }
      }
    } catch {}
  }
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
