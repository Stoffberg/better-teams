import { contextBridge, ipcRenderer } from "electron";

function invoke<T>(channel: string, ...args: unknown[]): Promise<T> {
  return ipcRenderer.invoke(channel, ...args) as Promise<T>;
}

contextBridge.exposeInMainWorld("betterTeams", {
  teams: {
    extractTokens: () => invoke("teams:extractTokens"),
    getAuthToken: (tenantId: string | null) =>
      invoke("teams:getAuthToken", tenantId),
    getAvailableAccounts: () => invoke("teams:getAvailableAccounts"),
    getCachedPresence: (userMris: string[]) =>
      invoke("teams:getCachedPresence", userMris),
  },
  images: {
    cacheImageFile: (
      cacheKey: string,
      bytes: Uint8Array,
      extension: string | null,
    ) => invoke("images:cacheFile", cacheKey, Array.from(bytes), extension),
    removeCachedImageFiles: (paths: string[]) =>
      invoke("images:removeFiles", paths),
    filePathToAssetUrl: (filePath: string) =>
      `better-teams-asset://file/${encodeURIComponent(filePath)}`,
  },
  sqlite: {
    execute: (sql: string, bindValues?: BetterTeamsBindValue[]) =>
      invoke("sqlite:execute", sql, bindValues ?? []),
    select: <T>(sql: string, bindValues?: BetterTeamsBindValue[]) =>
      invoke<T>("sqlite:select", sql, bindValues ?? []),
  },
  http: {
    fetch: (request: BetterTeamsFetchRequest) => invoke("http:fetch", request),
  },
  shell: {
    openExternal: (url: string) => invoke("shell:openExternal", url),
  },
} satisfies BetterTeamsDesktopApi);
