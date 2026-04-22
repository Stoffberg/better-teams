import { execFileSync } from "node:child_process";
import { createDecipheriv, pbkdf2Sync } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import Database from "better-sqlite3";

type ExtractedToken = {
  host: string;
  name: string;
  token: string;
  audience?: string;
  upn?: string;
  tenantId?: string;
  skypeId?: string;
  expiresAt: string;
};

type AccountOption = {
  upn?: string;
  tenantId?: string;
};

type CachedPresenceEntry = {
  mri: string;
  presence: {
    availability?: string;
    activity?: string;
  };
};

type CookieRow = {
  host_key: string;
  name: string;
  encrypted_value: Buffer;
};

const pbkdf2Salt = Buffer.from("saltysalt");
const encryptedPrefix = Buffer.from("v10");
const cdlWorkerPrefix =
  "https://cdl-worker/cdl-worker-cache-manager/cdl-worker-aggregated-user-data/";
const presenceCacheTtlMs = 30_000;

let presenceCache: {
  expiresAt: number;
  entries: CachedPresenceEntry[];
} | null = null;

export function extractTokens(): ExtractedToken[] {
  return extractAllTokens();
}

export function getAuthToken(tenantId?: string | null): ExtractedToken | null {
  return (
    extractAllTokens().find((token) => {
      if (token.name !== "authtoken") return false;
      return tenantId ? token.tenantId === tenantId : true;
    }) ?? null
  );
}

export function getAvailableAccounts(): AccountOption[] {
  const seen = new Set<string>();
  const accounts: AccountOption[] = [];

  for (const token of extractAllTokens()) {
    if (token.name !== "authtoken") continue;
    const key = `${token.upn ?? ""}\x1f${token.tenantId ?? ""}`;
    if (seen.has(key)) continue;
    seen.add(key);
    accounts.push({ upn: token.upn, tenantId: token.tenantId });
  }

  return accounts;
}

export function getCachedPresence(userMris: string[]): CachedPresenceEntry[] {
  const entries =
    presenceCache && presenceCache.expiresAt > Date.now()
      ? presenceCache.entries
      : scanAllPresence();

  presenceCache = {
    entries,
    expiresAt: Date.now() + presenceCacheTtlMs,
  };

  const byMri = new Map(
    entries.map((entry) => [entry.mri.trim().toLowerCase(), entry]),
  );

  return userMris
    .map((mri) => byMri.get(mri.trim().toLowerCase()))
    .filter((entry): entry is CachedPresenceEntry => Boolean(entry));
}

function extractAllTokens(): ExtractedToken[] {
  if (process.platform !== "darwin") {
    throw new Error(
      "Better Teams token extraction currently supports macOS Teams 2 only",
    );
  }

  const dbPath = cookiesDbPath();
  if (!fs.existsSync(dbPath)) {
    throw new Error("Teams cookies database not found");
  }

  const rows = readCookieRows(dbPath);
  const cachedKey = readCachedDecryptionKey();

  if (cachedKey) {
    const outcome = extractTokensWithKey(rows, cachedKey);
    if (outcome.tokens.length > 0 || outcome.decryptFailures === 0) {
      return outcome.tokens;
    }
  }

  const safeStorageKey = getSafeStorageKey();
  const decryptionKey = deriveDecryptionKey(safeStorageKey);
  const outcome = extractTokensWithKey(rows, decryptionKey);
  writeCachedDecryptionKey(decryptionKey);
  return outcome.tokens;
}

function cookiesDbPath(): string {
  return path.join(
    os.homedir(),
    "Library/Containers/com.microsoft.teams2/Data/Library/Application Support/Microsoft/MSTeams/EBWebView/WV2Profile_tfw/Cookies",
  );
}

function readCookieRows(dbPath: string): CookieRow[] {
  const db = new Database(dbPath, { readonly: true, fileMustExist: true });
  try {
    const rows = db
      .prepare(
        `SELECT host_key, name, encrypted_value
         FROM cookies
         WHERE (host_key LIKE '%teams%' OR host_key LIKE '%skype%')
           AND (name = 'authtoken' OR name = 'skypetoken_asm')
         ORDER BY expires_utc DESC`,
      )
      .all() as CookieRow[];

    return rows.map((row) => ({
      ...row,
      encrypted_value: Buffer.from(row.encrypted_value),
    }));
  } finally {
    db.close();
  }
}

function getSafeStorageKey(): Buffer {
  const password = execFileSync(
    "security",
    [
      "find-generic-password",
      "-w",
      "-s",
      "Microsoft Teams Safe Storage",
      "-a",
      "Microsoft Teams",
    ],
    { encoding: "utf8" },
  ).trimEnd();
  return Buffer.from(password, "utf8");
}

function cachedDecryptionKeyPath(): string {
  return path.join(
    os.homedir(),
    "Library/Application Support/Better Teams/teams-safe-storage-key.bin",
  );
}

function readCachedDecryptionKey(): Buffer | null {
  const filePath = cachedDecryptionKeyPath();
  if (!fs.existsSync(filePath)) return null;
  const key = fs.readFileSync(filePath);
  return key.length === 16 ? key : null;
}

function writeCachedDecryptionKey(key: Buffer): void {
  const filePath = cachedDecryptionKeyPath();
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, key, { mode: 0o600 });
  fs.chmodSync(filePath, 0o600);
}

function deriveDecryptionKey(safeStorageKey: Buffer): Buffer {
  return pbkdf2Sync(safeStorageKey, pbkdf2Salt, 1003, 16, "sha1");
}

function decryptValue(encrypted: Buffer, key: Buffer): string {
  if (encrypted.length === 0) return "";
  if (
    encrypted.length < 3 ||
    !encrypted.subarray(0, 3).equals(encryptedPrefix)
  ) {
    return encrypted.toString("utf8");
  }

  const iv = Buffer.alloc(16, 0x20);
  const decipher = createDecipheriv("aes-128-cbc", key, iv);
  const decrypted = Buffer.concat([
    decipher.update(encrypted.subarray(3)),
    decipher.final(),
  ]);
  return decrypted.toString("utf8");
}

function extractTokensWithKey(
  rows: CookieRow[],
  key: Buffer,
): { tokens: ExtractedToken[]; decryptFailures: number } {
  const now = Math.floor(Date.now() / 1000);
  const tokens: ExtractedToken[] = [];
  let decryptFailures = 0;

  for (const row of rows) {
    let decrypted = "";
    try {
      decrypted = decryptValue(row.encrypted_value, key);
    } catch {
      decryptFailures += 1;
      continue;
    }

    const jwt = extractJwt(decrypted);
    if (!jwt) continue;

    const payload = decodeJwtPayload(jwt);
    const exp = typeof payload?.exp === "number" ? payload.exp : 0;
    if (exp < now) continue;

    tokens.push({
      host: row.host_key,
      name: row.name,
      token: jwt,
      audience: stringValue(payload?.aud),
      upn: stringValue(payload?.upn),
      tenantId: stringValue(payload?.tid),
      skypeId: stringValue(payload?.skypeid),
      expiresAt: new Date(exp * 1000).toISOString(),
    });
  }

  tokens.sort((a, b) => b.expiresAt.localeCompare(a.expiresAt));
  return { tokens, decryptFailures };
}

function extractJwt(raw: string): string {
  const jwtPart = "eyJ[A-Za-z0-9_-]+\\.eyJ[A-Za-z0-9_-]+\\.[A-Za-z0-9_-]+";
  for (const pattern of [
    `Bearer%3D(${jwtPart})`,
    `Bearer%20(${jwtPart})`,
    `(${jwtPart})`,
  ]) {
    const match = raw.match(new RegExp(pattern));
    if (match?.[1]) return match[1];
  }
  return "";
}

function decodeJwtPayload(token: string): Record<string, unknown> | null {
  const payload = token.split(".")[1];
  if (!payload) return null;
  try {
    return JSON.parse(
      Buffer.from(payload, "base64url").toString("utf8"),
    ) as Record<string, unknown>;
  } catch {
    return null;
  }
}

function stringValue(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

function serviceWorkerCachePath(): string {
  return path.join(
    os.homedir(),
    "Library/Containers/com.microsoft.teams2/Data/Library/Application Support/Microsoft/MSTeams/EBWebView/WV2Profile_tfw/Service Worker/CacheStorage",
  );
}

function scanAllPresence(): CachedPresenceEntry[] {
  const cachePath = serviceWorkerCachePath();
  if (!fs.existsSync(cachePath)) return [];

  const entries = new Map<string, CachedPresenceEntry>();
  for (const filePath of collectFiles(cachePath)) {
    let bytes: Buffer;
    try {
      bytes = fs.readFileSync(filePath);
    } catch {
      continue;
    }
    const entry = parseCachedPresenceEntry(bytes);
    if (!entry) continue;
    entries.set(entry.mri.trim().toLowerCase(), entry);
  }
  return [...entries.values()];
}

function collectFiles(dirPath: string): string[] {
  const found: string[] = [];
  for (const entry of fs.readdirSync(dirPath, { withFileTypes: true })) {
    const entryPath = path.join(dirPath, entry.name);
    if (entry.isDirectory()) {
      found.push(...collectFiles(entryPath));
    } else if (entry.isFile()) {
      found.push(entryPath);
    }
  }
  return found;
}

function parseCachedPresenceEntry(bytes: Buffer): CachedPresenceEntry | null {
  const markerIndex = bytes.indexOf(cdlWorkerPrefix);
  if (markerIndex < 0) return null;
  const jsonStart = bytes.indexOf("{", markerIndex + cdlWorkerPrefix.length);
  if (jsonStart < 0) return null;
  const jsonBytes = extractJsonObject(bytes, jsonStart);
  if (!jsonBytes) return null;

  try {
    const envelope = JSON.parse(jsonBytes.toString("utf8")) as {
      presence?: {
        mri?: string;
        presence?: {
          availability?: string;
          activity?: string;
        };
      };
    };
    const mri = envelope.presence?.mri?.trim();
    const presence = envelope.presence?.presence;
    const availability = presence?.availability?.trim();
    const activity = presence?.activity?.trim();
    if (!mri || (!availability && !activity)) return null;
    return { mri, presence: { availability, activity } };
  } catch {
    return null;
  }
}

function extractJsonObject(bytes: Buffer, start: number): Buffer | null {
  if (bytes[start] !== 123) return null;

  let depth = 0;
  let inString = false;
  let escaped = false;

  for (let index = start; index < bytes.length; index += 1) {
    const byte = bytes[index];
    if (inString) {
      if (escaped) {
        escaped = false;
      } else if (byte === 92) {
        escaped = true;
      } else if (byte === 34) {
        inString = false;
      }
      continue;
    }

    if (byte === 34) {
      inString = true;
    } else if (byte === 123) {
      depth += 1;
    } else if (byte === 125) {
      depth -= 1;
      if (depth === 0) return bytes.subarray(start, index + 1);
    }
  }

  return null;
}
