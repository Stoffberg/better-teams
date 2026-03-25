/**
 * Debug: find ALL conversations with image attachments and test fetching.
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

  // Get ALL conversations
  const convUrl = `${chatService}/v1/users/ME/conversations?view=msnp24Equivalent|supportsMessageProperties&pageSize=100&startTime=0&targetType=Passport|Skype|Lync|Thread|NotificationStream|cnsTopicService|Agent`;
  const convData = await (
    await fetch(convUrl, {
      headers: { Authentication: `skypetoken=${skypeToken}` },
    })
  ).json();

  const convos = convData.conversations ?? [];
  console.log(`Total conversations: ${convos.length}`);

  // Check last messages for any with Picture type URIObject
  let foundImage = false;
  for (const conv of convos) {
    const content = conv.lastMessage?.content || "";
    if (
      content.includes('type="Picture') ||
      content.includes("type='Picture")
    ) {
      console.log(`\nFound image in conversation: ${conv.id}`);
      console.log(`  Last sender: ${conv.lastMessage?.imdisplayname}`);
      console.log(`  Content snippet: ${content.slice(0, 200)}`);
      foundImage = true;

      // Fetch messages from this conversation
      const encoded = encodeURIComponent(conv.id);
      const msgUrl = `${chatService}/v1/users/ME/conversations/${encoded}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=30&startTime=1`;
      const msgData = await (
        await fetch(msgUrl, {
          headers: { Authentication: `skypetoken=${skypeToken}` },
        })
      ).json();

      for (const m of msgData.messages ?? []) {
        const mc = m.content || "";
        if (
          mc.includes("URIObject") &&
          (mc.includes('type="Picture') || mc.includes("type='Picture"))
        ) {
          const uriMatch = mc.match(/uri="([^"]+)"/i);
          const thumbMatch = mc.match(/url_thumbnail="([^"]+)"/i);
          const nameMatch = mc.match(/OriginalName[^>]*v="([^"]+)"/i);
          console.log(`\n  Image message: ${m.id} (${m.composetime})`);
          console.log(`    From: ${m.imdisplayname}`);
          console.log(`    URI: ${uriMatch?.[1]}`);
          console.log(`    Thumbnail: ${thumbMatch?.[1]}`);
          console.log(`    OriginalName: ${nameMatch?.[1]}`);

          if (uriMatch?.[1]) {
            const uri = uriMatch[1];
            // Try different auth methods
            for (const [label, headers] of [
              ["skypetoken", { Authentication: `skypetoken=${skypeToken}` }],
              ["bearer", { Authorization: `Bearer ${authToken}` }],
              [
                "both",
                {
                  Authentication: `skypetoken=${skypeToken}`,
                  Authorization: `Bearer ${authToken}`,
                },
              ],
            ]) {
              const viewUrl = `${uri}/views/imgpsh_fullsize_anim`;
              try {
                const res = await fetch(viewUrl, {
                  headers,
                  redirect: "follow",
                });
                console.log(
                  `    [${label}] fullsize → ${res.status} ${res.statusText} (ct: ${res.headers.get("content-type")}, len: ${res.headers.get("content-length")})`,
                );
                if (res.ok) break;
              } catch (e) {
                console.log(`    [${label}] fullsize → Error: ${e.message}`);
              }
            }

            if (thumbMatch?.[1]) {
              for (const [label, headers] of [
                ["skypetoken", { Authentication: `skypetoken=${skypeToken}` }],
                ["bearer", { Authorization: `Bearer ${authToken}` }],
              ]) {
                try {
                  const res = await fetch(thumbMatch[1], {
                    headers,
                    redirect: "follow",
                  });
                  console.log(
                    `    [${label}] thumbnail → ${res.status} ${res.statusText} (ct: ${res.headers.get("content-type")}, len: ${res.headers.get("content-length")})`,
                  );
                  if (res.ok) break;
                } catch (e) {
                  console.log(`    [${label}] thumbnail → Error: ${e.message}`);
                }
              }
            }
          }
        }
      }
    }
  }

  if (!foundImage) {
    console.log(
      "\nNo Picture-type URIObject found in any conversation's last message.",
    );
    console.log("Scanning first 5 conversations for image messages...");

    for (const conv of convos.slice(0, 5)) {
      const encoded = encodeURIComponent(conv.id);
      const msgUrl = `${chatService}/v1/users/ME/conversations/${encoded}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=30&startTime=1`;
      try {
        const msgData = await (
          await fetch(msgUrl, {
            headers: { Authentication: `skypetoken=${skypeToken}` },
          })
        ).json();

        for (const m of msgData.messages ?? []) {
          const mc = m.content || "";
          if (mc.includes("URIObject") && mc.includes('type="Picture')) {
            const uriMatch = mc.match(/uri="([^"]+)"/i);
            const thumbMatch = mc.match(/url_thumbnail="([^"]+)"/i);
            console.log(`\n  Found image in conv ${conv.id}:`);
            console.log(
              `    Message: ${m.id} from ${m.imdisplayname} at ${m.composetime}`,
            );
            console.log(`    URI: ${uriMatch?.[1]}`);
            console.log(`    Thumbnail: ${thumbMatch?.[1]}`);

            if (uriMatch?.[1]) {
              for (const [label, headers] of [
                ["skypetoken", { Authentication: `skypetoken=${skypeToken}` }],
                ["bearer", { Authorization: `Bearer ${authToken}` }],
                [
                  "both",
                  {
                    Authentication: `skypetoken=${skypeToken}`,
                    Authorization: `Bearer ${authToken}`,
                  },
                ],
              ]) {
                const viewUrl = `${uriMatch[1]}/views/imgpsh_fullsize_anim`;
                try {
                  const res = await fetch(viewUrl, {
                    headers,
                    redirect: "follow",
                  });
                  console.log(
                    `    [${label}] → ${res.status} (ct: ${res.headers.get("content-type")}, len: ${res.headers.get("content-length")})`,
                  );
                  if (res.ok) {
                    foundImage = true;
                    break;
                  }
                } catch (e) {
                  console.log(`    [${label}] → Error: ${e.message}`);
                }
              }
            }
            if (foundImage) break;
          }
        }
      } catch {}
      if (foundImage) break;
    }
  }
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
