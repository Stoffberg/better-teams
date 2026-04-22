import { z } from "zod";

export const RawTokenSchema = z.object({
  host: z.string(),
  name: z.string(),
  token: z.string(),
  audience: z.string().optional(),
  upn: z.string().optional(),
  tenantId: z.string().optional(),
  skypeId: z.string().optional(),
  expiresAt: z.string(),
});

export const PresenceInfoSchema = z.object({
  availability: z.string().optional(),
  activity: z.string().optional(),
});

export const PresenceEntrySchema = z.object({
  mri: z.string(),
  presence: PresenceInfoSchema,
});

export const AccountOptionSchema = z.object({
  upn: z.string().optional(),
  tenantId: z.string().optional(),
});

export const TeamsSessionInfoSchema = z.object({
  upn: z.string().optional(),
  tenantId: z.string(),
  skypeId: z.string().optional(),
  expiresAt: z.string().nullable(),
  region: z.string().nullable(),
});

export const TenantIdSchema = z.string().nullable();

export const PresenceRequestSchema = z.array(z.string());

export const ProfilePresentationRequestSchema = z.array(z.string());

export const CachedConversationSchema = z
  .object({ id: z.string() })
  .passthrough();

export const CachedProfilePresentationSchema = z.object({
  avatarThumbs: z.record(z.string(), z.string()),
  avatarFull: z.record(z.string(), z.string()),
  displayNames: z.record(z.string(), z.string()),
  emails: z.record(z.string(), z.string()),
  jobTitles: z.record(z.string(), z.string()),
  departments: z.record(z.string(), z.string()),
  companyNames: z.record(z.string(), z.string()),
  tenantNames: z.record(z.string(), z.string()),
  locations: z.record(z.string(), z.string()),
});

export const CachedMessagesResponseSchema = z
  .object({
    messages: z.array(z.object({ id: z.string() }).passthrough()).optional(),
  })
  .passthrough();

export const ImageCacheRequestSchema = z.object({
  cacheKey: z.string(),
  bytes: z.instanceof(Uint8Array),
  extension: z.string().nullable(),
});

export const ImageCacheIpcRequestSchema = ImageCacheRequestSchema.extend({
  bytes: z.array(z.number().int().min(0).max(255)),
});

export const ImageCachePathSchema = z.string();

export const ShellOpenExternalUrlSchema = z
  .string()
  .url()
  .refine((value) =>
    ["https:", "http:", "mailto:"].includes(new URL(value).protocol),
  );

export const FetchRequestSchema = z.object({
  url: z.string().url(),
  method: z.string().optional(),
  headers: z.array(z.tuple([z.string(), z.string()])),
  body: z.union([z.string(), z.instanceof(ArrayBuffer), z.null()]),
});

export const FetchResponseSchema = z.object({
  status: z.number(),
  statusText: z.string(),
  headers: z.array(z.tuple([z.string(), z.string()])),
  body: z.instanceof(ArrayBuffer),
});

export type BetterTeamsRawToken = z.infer<typeof RawTokenSchema>;
export type BetterTeamsPresenceInfo = z.infer<typeof PresenceInfoSchema>;
export type BetterTeamsPresenceEntry = z.infer<typeof PresenceEntrySchema>;
export type BetterTeamsAccountOption = z.infer<typeof AccountOptionSchema>;
export type BetterTeamsSessionInfo = z.infer<typeof TeamsSessionInfoSchema>;
export type BetterTeamsCachedConversation = z.infer<
  typeof CachedConversationSchema
>;
export type BetterTeamsCachedProfilePresentation = z.infer<
  typeof CachedProfilePresentationSchema
>;
export type BetterTeamsCachedMessagesResponse = z.infer<
  typeof CachedMessagesResponseSchema
>;
export type BetterTeamsImageCacheRequest = z.infer<
  typeof ImageCacheRequestSchema
>;
export type BetterTeamsFetchRequest = z.infer<typeof FetchRequestSchema>;
export type BetterTeamsFetchResponse = z.infer<typeof FetchResponseSchema>;

export interface BetterTeamsDesktopApi {
  teams: {
    extractTokens(): Promise<BetterTeamsRawToken[]>;
    getAuthToken(tenantId: string | null): Promise<BetterTeamsRawToken | null>;
    getAvailableAccounts(): Promise<BetterTeamsAccountOption[]>;
    getCachedAccounts(): Promise<BetterTeamsAccountOption[]>;
    getCachedSession(
      tenantId: string | null,
    ): Promise<BetterTeamsSessionInfo | null>;
    getCachedPresence(userMris: string[]): Promise<BetterTeamsPresenceEntry[]>;
    getCachedProfilePresentation(
      mris: string[],
    ): Promise<BetterTeamsCachedProfilePresentation>;
    getCachedConversations(
      tenantId: string | null,
    ): Promise<BetterTeamsCachedConversation[]>;
    getCachedMessages(
      tenantId: string | null,
      conversationId: string,
    ): Promise<BetterTeamsCachedMessagesResponse | null>;
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
