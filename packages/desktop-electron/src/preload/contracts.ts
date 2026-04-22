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

export const TenantIdSchema = z.string().nullable();

export const PresenceRequestSchema = z.array(z.string());

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
