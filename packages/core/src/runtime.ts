import type { ExtractedToken, PresenceInfo } from "./teams/types";

export type TeamsRuntimeFetch = typeof globalThis.fetch;

export type TeamsRuntimePresenceEntry = {
  mri: string;
  presence: PresenceInfo;
};

export type TeamsRuntimeAccountOption = {
  upn?: string;
  tenantId?: string;
};

export type TeamsRuntime = {
  fetch: TeamsRuntimeFetch;
  extractTokens: () => Promise<ExtractedToken[]>;
  getAuthToken: (tenantId?: string | null) => Promise<ExtractedToken | null>;
  getAvailableAccounts: () => Promise<TeamsRuntimeAccountOption[]>;
  getCachedPresence: (
    userMris: string[],
  ) => Promise<TeamsRuntimePresenceEntry[]>;
  cacheImageFile: (
    cacheKey: string,
    bytes: Uint8Array,
    extension?: string | null,
  ) => Promise<string>;
  getCachedImageFile: (cacheKey: string) => Promise<string | null>;
  hasCachedImageFile: (filePath: string) => Promise<boolean>;
  filePathToAssetUrl: (filePath: string) => string;
};

let runtime: TeamsRuntime | null = null;

export function registerTeamsRuntime(next: TeamsRuntime): void {
  runtime = next;
}

export function getTeamsRuntime(): TeamsRuntime {
  if (!runtime) {
    throw new Error("Better Teams runtime has not been registered");
  }
  return runtime;
}
