import {
  registerTeamsRuntime,
  type TeamsRuntime,
} from "@better-teams/core/runtime";
import type {
  Conversation,
  ExtractedToken,
  MessagesResponse,
  PresenceInfo,
  TeamsProfilePresentation,
  TeamsSessionInfo,
} from "@better-teams/core/teams/types";
import type { BetterTeamsDesktopApi } from "@better-teams/desktop-electron/preload";
import { fetch as desktopFetch } from "./fetch";

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
  if (Array.isArray(accounts) && accounts.length > 0) return accounts;
  const cachedAccounts = await getCachedAccounts();
  if (Array.isArray(cachedAccounts) && cachedAccounts.length > 0) {
    return cachedAccounts;
  }
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

export async function getCachedAccounts() {
  const accounts = await api().teams.getCachedAccounts();
  return Array.isArray(accounts) ? accounts : [];
}

export async function getCachedSession(
  tenantId?: string | null,
): Promise<TeamsSessionInfo | null> {
  return api().teams.getCachedSession(
    tenantId ?? null,
  ) as Promise<TeamsSessionInfo | null>;
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

export async function getCachedConversations(
  tenantId?: string | null,
): Promise<Conversation[]> {
  return api().teams.getCachedConversations(tenantId ?? null) as Promise<
    Conversation[]
  >;
}

export async function getCachedMessages(
  tenantId: string | null | undefined,
  conversationId: string,
): Promise<MessagesResponse | null> {
  return api().teams.getCachedMessages(
    tenantId ?? null,
    conversationId,
  ) as Promise<MessagesResponse | null>;
}

export async function getCachedProfilePresentation(
  mris: string[],
): Promise<TeamsProfilePresentation> {
  return api().teams.getCachedProfilePresentation(
    mris,
  ) as Promise<TeamsProfilePresentation>;
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
