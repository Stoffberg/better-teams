import { contextBridge, ipcRenderer } from "electron";
import { z } from "zod";
import {
  AccountOptionSchema,
  type BetterTeamsDesktopApi,
  type BetterTeamsFetchRequest,
  CachedConversationSchema,
  CachedMessagesResponseSchema,
  CachedProfilePresentationSchema,
  FetchRequestSchema,
  FetchResponseSchema,
  ImageCachePathSchema,
  ImageCacheRequestSchema,
  PresenceEntrySchema,
  PresenceRequestSchema,
  ProfilePresentationRequestSchema,
  RawTokenSchema,
  ShellOpenExternalUrlSchema,
  TeamsSessionInfoSchema,
  TenantIdSchema,
} from "./contracts";

function invoke<T>(channel: string, ...args: unknown[]): Promise<T> {
  return ipcRenderer.invoke(channel, ...args) as Promise<T>;
}

contextBridge.exposeInMainWorld("betterTeams", {
  teams: {
    extractTokens: async () =>
      RawTokenSchema.array().parse(await invoke("teams:extractTokens")),
    getAuthToken: async (tenantId: string | null) =>
      RawTokenSchema.nullable().parse(
        await invoke("teams:getAuthToken", TenantIdSchema.parse(tenantId)),
      ),
    getAvailableAccounts: async () =>
      AccountOptionSchema.array().parse(
        await invoke("teams:getAvailableAccounts"),
      ),
    getCachedAccounts: async () =>
      AccountOptionSchema.array().parse(
        await invoke("teams:getCachedAccounts"),
      ),
    getCachedSession: async (tenantId: string | null) =>
      TeamsSessionInfoSchema.nullable().parse(
        await invoke("teams:getCachedSession", TenantIdSchema.parse(tenantId)),
      ),
    getCachedPresence: async (userMris: string[]) =>
      PresenceEntrySchema.array().parse(
        await invoke(
          "teams:getCachedPresence",
          PresenceRequestSchema.parse(userMris),
        ),
      ),
    getCachedProfilePresentation: async (mris: string[]) =>
      CachedProfilePresentationSchema.parse(
        await invoke(
          "teams:getCachedProfilePresentation",
          ProfilePresentationRequestSchema.parse(mris),
        ),
      ),
    getCachedConversations: async (tenantId: string | null) =>
      CachedConversationSchema.array().parse(
        await invoke(
          "teams:getCachedConversations",
          TenantIdSchema.parse(tenantId),
        ),
      ),
    getCachedMessages: async (
      tenantId: string | null,
      conversationId: string,
    ) =>
      CachedMessagesResponseSchema.nullable().parse(
        await invoke(
          "teams:getCachedMessages",
          TenantIdSchema.parse(tenantId),
          z.string().parse(conversationId),
        ),
      ),
  },
  images: {
    cacheImageFile: async (
      cacheKey: string,
      bytes: Uint8Array,
      extension: string | null,
    ) => {
      const request = ImageCacheRequestSchema.parse({
        cacheKey,
        bytes,
        extension,
      });
      return ImageCachePathSchema.parse(
        await invoke(
          "images:cacheFile",
          request.cacheKey,
          Array.from(request.bytes),
          request.extension,
        ),
      );
    },
    getCachedImageFile: async (cacheKey: string) =>
      ImageCachePathSchema.nullable().parse(
        await invoke(
          "images:getCachedFile",
          ImageCachePathSchema.parse(cacheKey),
        ),
      ),
    hasCachedImageFile: async (filePath: string) =>
      z
        .boolean()
        .parse(
          await invoke(
            "images:hasCachedFile",
            ImageCachePathSchema.parse(filePath),
          ),
        ),
    filePathToAssetUrl: (filePath: string) =>
      `better-teams-asset://file/${encodeURIComponent(
        ImageCachePathSchema.parse(filePath),
      )}`,
  },
  http: {
    fetch: async (request: BetterTeamsFetchRequest) =>
      FetchResponseSchema.parse(
        await invoke("http:fetch", FetchRequestSchema.parse(request)),
      ),
  },
  shell: {
    openExternal: async (url: string) => {
      await invoke("shell:openExternal", ShellOpenExternalUrlSchema.parse(url));
    },
  },
} satisfies BetterTeamsDesktopApi);
