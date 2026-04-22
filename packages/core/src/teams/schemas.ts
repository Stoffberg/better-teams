import { z } from "zod";
import {
  conversationLastActivityTime,
  conversationMemberCount,
  extractTeamsMri,
  messageSenderDisplayName,
  normalizeConversationProperties,
  normalizeMessageProperties,
  normalizeTeamsTimestamp,
  specialThreadTypeFromConversationId,
} from "./normalize";
import type {
  AuthzResponse,
  Conversation,
  ConversationsResponse,
  MembersConsumptionHorizonsResponse,
  Message,
  MessagesResponse,
  PresenceInfo,
  TeamsProfilePresentation,
} from "./types";

const MetadataSchema = z
  .object({
    backwardLink: z.string().optional(),
    syncState: z.string().optional(),
    lastCompleteSegmentStartTime: z.number().optional(),
    lastCompleteSegmentEndTime: z.number().optional(),
  })
  .passthrough();

const skypeWireString = z
  .union([
    z.string(),
    z.number(),
    z.boolean(),
    z.bigint(),
    z.null(),
    z.undefined(),
  ])
  .transform((v) => (v == null ? "" : String(v)));

const optionalConversationIdWire = z
  .union([z.string(), z.number(), z.bigint(), z.null()])
  .transform((v) => (v == null ? undefined : String(v)))
  .optional();

const MessageWireSchema = z
  .object({
    id: skypeWireString,
    sequenceId: z.number().optional(),
    clientMessageId: z.string().optional(),
    clientmessageid: skypeWireString.optional(),
    conversationId: optionalConversationIdWire,
    conversationid: optionalConversationIdWire,
    type: skypeWireString,
    messagetype: skypeWireString,
    contenttype: skypeWireString,
    content: z.string().optional(),
    from: skypeWireString,
    fromTenantId: z.string().optional(),
    imdisplayname: z.string().optional(),
    conversationLink: z.string().optional(),
    version: skypeWireString.optional(),
    postType: z.string().optional(),
    s2spartnername: z.string().optional(),
    skypeguid: z.string().optional(),
    rootMessageId: skypeWireString.optional(),
    clumpId: z.string().optional(),
    secondaryReferenceId: z.string().optional(),
    prioritizeImDisplayName: z.boolean().optional(),
    fromDisplayNameInToken: z.string().optional(),
    fromGivenNameInToken: z.string().optional(),
    fromFamilyNameInToken: z.string().optional(),
    amsreferences: z.array(z.unknown()).optional(),
    annotationsSummary: z.record(z.string(), z.unknown()).optional(),
    composetime: skypeWireString,
    originalarrivaltime: skypeWireString,
    properties: z.record(z.string(), z.unknown()).optional(),
  })
  .passthrough();

export const MessageSchema: z.ZodType<Message> = MessageWireSchema.transform(
  (m, ctx) => {
    const conversationId = m.conversationId ?? m.conversationid;
    if (!conversationId) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "Missing conversationId",
      });
      return z.NEVER;
    }
    const composetime = normalizeTeamsTimestamp(m.composetime);
    const originalarrivaltime = normalizeTeamsTimestamp(m.originalarrivaltime);
    const timestamp = originalarrivaltime || composetime;
    const properties = normalizeMessageProperties(m.properties);
    return {
      ...m,
      clientMessageId: m.clientMessageId ?? m.clientmessageid,
      conversationId,
      fromMri: extractTeamsMri(m.from),
      senderDisplayName: messageSenderDisplayName(m),
      composetime,
      originalarrivaltime,
      timestamp,
      version: m.version,
      rootMessageId: m.rootMessageId,
      properties,
      deleted: Boolean(properties?.deletetime || properties?.hardDeleteTime),
    };
  },
);

function lastMessageWireWithParentConversationId(
  parentConversationId: string,
  raw: unknown,
): unknown {
  if (raw == null) return undefined;
  if (typeof raw !== "object" || Array.isArray(raw)) return raw;
  const o = { ...(raw as Record<string, unknown>) };
  const has = (v: unknown) => typeof v === "string" && v.trim().length > 0;
  if (!has(o.conversationId) && !has(o.conversationid)) {
    o.conversationId = parentConversationId;
  }
  return o;
}

const ConversationMemberSchema = z
  .object({
    id: z.string(),
    role: z.string().optional().default("User"),
    isMri: z.boolean().optional().default(true),
    displayName: z.string().optional(),
    friendlyName: z.string().optional(),
    tenantId: z.string().optional(),
    objectId: z.string().optional(),
    userPrincipalName: z.string().optional(),
  })
  .passthrough();

const MemberConsumptionHorizonSchema = z
  .object({
    id: z.string(),
    consumptionhorizon: z.string(),
    messageVisibilityTime: z.number().optional(),
  })
  .passthrough();

export const MembersConsumptionHorizonsResponseSchema: z.ZodType<MembersConsumptionHorizonsResponse> =
  z
    .object({
      id: z.string(),
      version: z.string().optional(),
      consumptionhorizons: z.array(MemberConsumptionHorizonSchema).default([]),
    })
    .passthrough();

const ThreadPropertiesSchema = z
  .object({
    topic: z.string().optional(),
    threadType: z.string().optional(),
    productThreadType: z.string().optional(),
    lastjoinat: z.string().optional(),
    membercount: z.string().optional(),
    memberCount: z.union([z.string(), z.number()]).optional(),
  })
  .passthrough();

const ConversationRowSchema = z
  .object({
    id: z.string(),
    conversationType: z.string().optional(),
    properties: z.record(z.string(), z.unknown()).optional(),
    threadProperties: ThreadPropertiesSchema.optional(),
    lastMessage: z.unknown().optional(),
    members: z.array(ConversationMemberSchema).optional(),
  })
  .passthrough();

export const ConversationSchema: z.ZodType<Conversation> =
  ConversationRowSchema.transform((c) => {
    const enriched = lastMessageWireWithParentConversationId(
      c.id,
      c.lastMessage,
    );
    const lastMessage =
      enriched == null ? undefined : MessageSchema.parse(enriched);
    const properties = normalizeConversationProperties(
      c.properties as Record<string, unknown> | undefined,
    );
    const memberCount = conversationMemberCount({
      ...c,
      properties,
      lastMessage,
    });
    return {
      ...c,
      properties,
      lastMessage,
      memberCount,
      consumptionHorizon:
        typeof properties?.consumptionhorizon === "string"
          ? properties.consumptionhorizon
          : undefined,
      lastActivityTime: lastMessage
        ? conversationLastActivityTime({ lastMessage })
        : "",
      specialThreadType: specialThreadTypeFromConversationId(c.id),
    };
  });

export const ConversationsResponseSchema: z.ZodType<ConversationsResponse> = z
  .object({
    conversations: z.array(ConversationSchema),
    _metadata: MetadataSchema.optional(),
  })
  .passthrough();

export const MessagesResponseSchema: z.ZodType<MessagesResponse> = z
  .object({
    messages: z.array(MessageSchema),
    _metadata: MetadataSchema.optional(),
  })
  .passthrough();

const PresenceInfoSchema: z.ZodType<PresenceInfo> = z
  .object({
    availability: z.string().optional(),
    activity: z.string().optional(),
    statusMessage: z
      .object({
        message: z.string(),
        expiry: z.string(),
      })
      .passthrough()
      .optional(),
  })
  .passthrough();

export const PresenceEntrySchema = z
  .object({
    mri: z.string().optional(),
    presence: PresenceInfoSchema.optional(),
    availability: z.string().optional(),
    activity: z.string().optional(),
  })
  .passthrough()
  .transform((entry) => ({
    mri: entry.mri ?? "",
    presence:
      entry.presence ??
      ({
        availability: entry.availability,
        activity: entry.activity,
      } satisfies PresenceInfo),
  }));

export const PresenceListSchema = z.array(PresenceEntrySchema);

const AuthzTokensSchema = z
  .object({
    skypeToken: z.string(),
    expiresIn: z.number().catch(0),
  })
  .passthrough();

export const AuthzResponseSchema: z.ZodType<AuthzResponse> = z
  .object({
    tokens: AuthzTokensSchema,
    region: z.string(),
    regionGtms: z
      .object({
        chatService: z.string(),
        chatServiceAggregator: z.string(),
        middleTier: z.string(),
        unifiedPresence: z.string(),
        ams: z.string(),
        amsV2: z.string(),
      })
      .catchall(z.string()),
    regionSettings: z.record(z.string(), z.boolean()).default({}),
    licenseDetails: z
      .object({
        isFreemium: z.boolean().default(false),
        isTrial: z.boolean().default(false),
      })
      .passthrough()
      .default({ isFreemium: false, isTrial: false }),
  })
  .passthrough();

export const ShortProfileRowSchema = z
  .object({
    mri: z.string().optional(),
    userMri: z.string().optional(),
    objectId: z.string().optional(),
    homeMri: z.string().optional(),
    homeTenantId: z.string().optional(),
    userPrincipalName: z.string().optional(),
    givenName: z.string().optional(),
    surname: z.string().optional(),
    displayName: z.string().optional(),
    displayname: z.string().optional(),
    name: z.string().optional(),
    email: z.string().optional(),
    mail: z.string().optional(),
    jobTitle: z.string().optional(),
    title: z.string().optional(),
    department: z.string().optional(),
    userLocation: z.string().optional(),
    companyName: z.string().optional(),
    userType: z.string().optional(),
    userState: z.string().optional(),
    tenantName: z.string().optional(),
    isShortProfile: z.boolean().optional(),
    imageUri: z.string().optional(),
    imageURL: z.string().optional(),
    profileImageUri: z.string().optional(),
    profilePictureUri: z.string().optional(),
    profileImageUrl: z.string().optional(),
    avatarUrl: z.string().optional(),
    avatarURL: z.string().optional(),
    pictureUrl: z.string().optional(),
    highResolutionImageUrl: z.string().optional(),
    linkedInProfilePictureUrl: z.string().optional(),
    skypeTeamsInfo: z.record(z.string(), z.unknown()).optional(),
    shortProfile: z.record(z.string(), z.unknown()).optional(),
  })
  .passthrough();

const ShortProfileEnvelopeSchema = z.union([
  z.array(ShortProfileRowSchema),
  z.object({ value: z.array(ShortProfileRowSchema) }).passthrough(),
  z.object({ profiles: z.array(ShortProfileRowSchema) }).passthrough(),
  z.object({ users: z.array(ShortProfileRowSchema) }).passthrough(),
  z.object({ shortProfiles: z.array(ShortProfileRowSchema) }).passthrough(),
  z.record(z.string(), ShortProfileRowSchema),
]);

export function parseShortProfileRows(raw: unknown): ShortProfileRow[] {
  const parsed = ShortProfileEnvelopeSchema.parse(raw);
  if (Array.isArray(parsed)) return parsed;
  if ("value" in parsed && Array.isArray(parsed.value)) return parsed.value;
  if ("profiles" in parsed && Array.isArray(parsed.profiles))
    return parsed.profiles;
  if ("users" in parsed && Array.isArray(parsed.users)) return parsed.users;
  if ("shortProfiles" in parsed && Array.isArray(parsed.shortProfiles)) {
    return parsed.shortProfiles;
  }
  return Object.entries(parsed).map(([mri, row]) => {
    const withMri = ShortProfileRowSchema.parse(row);
    if (typeof withMri.mri !== "string" || !withMri.mri.startsWith("8:")) {
      withMri.mri = mri;
    }
    return withMri;
  });
}

export const UserProfileResponseSchema = z.union([
  z.object({ value: z.record(z.string(), z.unknown()) }).passthrough(),
  z.record(z.string(), z.unknown()),
]);

export const TeamsProfilePresentationSchema: z.ZodType<TeamsProfilePresentation> =
  z.object({
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

export type ShortProfileRow = z.infer<typeof ShortProfileRowSchema>;
