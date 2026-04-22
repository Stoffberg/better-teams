/// <reference types="vite/client" />

type BetterTeamsBindValue = string | number | boolean | null;

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
    removeCachedImageFiles(paths: string[]): Promise<void>;
    filePathToAssetUrl(filePath: string): string;
  };
  sqlite: {
    execute(sql: string, bindValues?: BetterTeamsBindValue[]): Promise<void>;
    select<T>(sql: string, bindValues?: BetterTeamsBindValue[]): Promise<T>;
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
