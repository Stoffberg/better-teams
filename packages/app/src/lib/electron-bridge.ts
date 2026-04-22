import {
  registerTeamsRuntime,
  type TeamsRuntime,
} from "@better-teams/core/runtime";
import type {
  ExtractedToken,
  PresenceInfo,
} from "@better-teams/core/teams/types";
import type { BetterTeamsDesktopApi } from "@better-teams/desktop-electron/preload";
import { fetch as desktopFetch } from "./electron-fetch";

const PRESENCE_CACHE_TTL_MS = 15_000;
const cachedPresenceByMri = new Map<
  string,
  { expiresAt: number; presence: PresenceInfo }
>();
const cachedPresenceRequests = new Map<
  string,
  Promise<Record<string, PresenceInfo>>
>();

export async function extractTokens(): Promise<ExtractedToken[]> {
  const raw = await api().teams.extractTokens();
  return raw.map(hydrateToken);
}

export async function getAuthToken(
  tenantId?: string | null,
): Promise<ExtractedToken | null> {
  const raw = await api().teams.getAuthToken(tenantId ?? null);
  return raw ? hydrateToken(raw) : null;
}

export async function getAvailableAccounts() {
  const accounts = await api().teams.getAvailableAccounts();
  if (accounts.length > 0) return accounts;
  const tokens = await extractTokens();
  const byTenant = new Map<string, { upn?: string; tenantId?: string }>();
  for (const token of tokens) {
    const key = token.tenantId ?? token.upn;
    if (!key) continue;
    byTenant.set(key, {
      upn: token.upn,
      tenantId: token.tenantId,
    });
  }
  return [...byTenant.values()];
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
    request = api()
      .teams.getCachedPresence(missingMris)
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
  extension?: string | null,
): Promise<string> {
  return api().images.cacheImageFile(cacheKey, bytes, extension ?? null);
}

export async function getCachedImageFile(
  cacheKey: string,
): Promise<string | null> {
  return api().images.getCachedImageFile(cacheKey);
}

export async function hasCachedImageFile(filePath: string): Promise<boolean> {
  return api().images.hasCachedImageFile(filePath);
}

export function filePathToAssetUrl(filePath: string): string {
  return api().images.filePathToAssetUrl(filePath);
}

type RawExtractedToken = Omit<ExtractedToken, "expiresAt"> & {
  expiresAt: string;
};

function hydrateToken(raw: RawExtractedToken): ExtractedToken {
  return {
    ...raw,
    expiresAt: new Date(raw.expiresAt),
  };
}

function api(): BetterTeamsDesktopApi {
  if (!window.betterTeams) {
    throw new Error("Better Teams desktop API is not available");
  }
  return window.betterTeams;
}

export function registerDesktopTeamsRuntime(): void {
  const runtime: TeamsRuntime = {
    fetch: desktopFetch,
    extractTokens,
    getAuthToken,
    getAvailableAccounts,
    getCachedPresence: (userMris) => api().teams.getCachedPresence(userMris),
    cacheImageFile,
    getCachedImageFile,
    hasCachedImageFile,
    filePathToAssetUrl,
  };
  registerTeamsRuntime(runtime);
}
