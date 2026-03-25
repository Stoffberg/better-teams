/**
 * Thin wrapper around Tauri `invoke()` for token extraction commands.
 *
 * These mirror the Rust commands in `src-tauri/src/commands/token.rs`
 * and replace the Electron IPC calls that previously went through
 * `window.betterTeams.teams.*`.
 */

import { convertFileSrc, invoke } from "@tauri-apps/api/core";
import type {
  ExtractedToken,
  PresenceInfo,
  TeamsAccountOption,
} from "@/services/teams/types";

const PRESENCE_CACHE_TTL_MS = 15_000;
const cachedPresenceByMri = new Map<
  string,
  { expiresAt: number; presence: PresenceInfo }
>();
const cachedPresenceRequests = new Map<
  string,
  Promise<Record<string, PresenceInfo>>
>();

/**
 * Extract all valid (non-expired) Teams auth tokens from the local
 * Teams cookie store, sorted by expiry descending.
 */
export async function extractTokens(): Promise<ExtractedToken[]> {
  const raw = await invoke<RawExtractedToken[]>("extract_tokens");
  return raw.map(hydrateToken);
}

/**
 * Get the best available auth token (Bearer token for api.spaces.skype.com).
 * Optionally filtered to a specific tenant.
 */
export async function getAuthToken(
  tenantId?: string,
): Promise<ExtractedToken | null> {
  const raw = await invoke<RawExtractedToken | null>("get_auth_token", {
    tenantId: tenantId ?? null,
  });
  return raw ? hydrateToken(raw) : null;
}

/**
 * Get all available Teams accounts (deduplicated across tenants).
 */
export async function getAvailableAccounts(): Promise<TeamsAccountOption[]> {
  return invoke<TeamsAccountOption[]>("get_available_accounts");
}

export async function getCachedPresence(
  userMris: string[],
): Promise<Record<string, PresenceInfo>> {
  if (userMris.length === 0) return {};

  const now = Date.now();
  const uniqueMris = [
    ...new Set(userMris.map((mri) => mri.trim()).filter(Boolean)),
  ];
  const cached: Record<string, PresenceInfo> = {};
  const missingMris: string[] = [];

  for (const mri of uniqueMris) {
    const entry = cachedPresenceByMri.get(mri);
    if (entry && entry.expiresAt > now) {
      cached[mri] = entry.presence;
      continue;
    }
    cachedPresenceByMri.delete(mri);
    missingMris.push(mri);
  }

  if (missingMris.length === 0) {
    return cached;
  }

  const requestKey = missingMris.slice().sort().join("\x1f");
  let request = cachedPresenceRequests.get(requestKey);
  if (!request) {
    request = invoke<RawCachedPresenceEntry[]>("get_cached_presence", {
      userMris: missingMris,
    })
      .then((entries) => {
        const fresh = Object.fromEntries(
          entries.map((entry) => [
            entry.mri,
            entry.presence satisfies PresenceInfo,
          ]),
        );
        const expiresAt = Date.now() + PRESENCE_CACHE_TTL_MS;
        for (const [mri, presence] of Object.entries(fresh)) {
          cachedPresenceByMri.set(mri, { expiresAt, presence });
        }
        return fresh;
      })
      .finally(() => {
        cachedPresenceRequests.delete(requestKey);
      });
    cachedPresenceRequests.set(requestKey, request);
  }

  return { ...cached, ...(await request) };
}

export async function cacheImageFile(
  cacheKey: string,
  bytes: Uint8Array,
  extension?: string,
): Promise<string> {
  return invoke<string>("cache_image_file", {
    cacheKey,
    bytes: Array.from(bytes),
    extension: extension ?? null,
  });
}

export async function removeCachedImageFiles(paths: string[]): Promise<void> {
  if (paths.length === 0) return;
  await invoke("remove_cached_image_files", { paths });
}

export function filePathToAssetUrl(filePath: string): string {
  return convertFileSrc(filePath);
}

// ── Internal ──

/** The token as it arrives from Rust (expiresAt is an ISO string). */
type RawExtractedToken = Omit<ExtractedToken, "expiresAt"> & {
  expiresAt: string;
};

type RawCachedPresenceEntry = {
  mri: string;
  presence: PresenceInfo;
};

/** Convert the ISO-string `expiresAt` from Rust into a JS `Date`. */
function hydrateToken(raw: RawExtractedToken): ExtractedToken {
  return {
    ...raw,
    expiresAt: new Date(raw.expiresAt),
  };
}
