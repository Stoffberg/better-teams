/// <reference types="vite/client" />

type BetterTeamsRawToken = {
  host: string;
  name: string;
  token: string;
  audience?: string;
  upn?: string;
  tenantId?: string;
  skypeId?: string;
  expiresAt: string;
};

type BetterTeamsPresenceInfo = {
  availability?: string;
  activity?: string;
};

type BetterTeamsPresenceEntry = {
  mri: string;
  presence: BetterTeamsPresenceInfo;
};

type BetterTeamsAccountOption = {
  upn?: string;
  tenantId?: string;
};

type BetterTeamsFetchRequest = {
  url: string;
  method?: string;
  headers: [string, string][];
  body: string | ArrayBuffer | null;
};

type BetterTeamsFetchResponse = {
  status: number;
  statusText: string;
  headers: [string, string][];
  body: ArrayBuffer;
};

interface BetterTeamsDesktopApi {
  teams: {
    extractTokens(): Promise<BetterTeamsRawToken[]>;
    getAuthToken(tenantId: string | null): Promise<BetterTeamsRawToken | null>;
    getAvailableAccounts(): Promise<BetterTeamsAccountOption[]>;
    getCachedPresence(userMris: string[]): Promise<BetterTeamsPresenceEntry[]>;
  };
  images: {
    cacheImageFile(
      cacheKey: string,
      bytes: Uint8Array,
      extension: string | null,
    ): Promise<string>;
    getCachedImageFile(cacheKey: string): Promise<string | null>;
    hasCachedImageFile(filePath: string): Promise<boolean>;
    filePathToAssetUrl(filePath: string): string;
  };
  http: {
    fetch(request: BetterTeamsFetchRequest): Promise<BetterTeamsFetchResponse>;
  };
  shell: {
    openExternal(url: string): Promise<void>;
  };
}

interface Window {
  betterTeams?: BetterTeamsDesktopApi;
}
