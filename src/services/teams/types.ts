export interface TeamsAccountOption {
  upn?: string;
  tenantId?: string;
}

export type TeamsTenantId = string;

export interface TeamsSessionInfo {
  upn?: string;
  tenantId: TeamsTenantId;
  skypeId?: string;
  expiresAt: string | null;
  region: string | null;
}

export interface AuthzResponse {
  tokens: {
    skypeToken: string;
    expiresIn: number;
  };
  region: string;
  regionGtms: RegionGtms;
  regionSettings: Record<string, boolean>;
  licenseDetails: {
    isFreemium: boolean;
    isTrial: boolean;
  };
}

export interface RegionGtms {
  chatService: string;
  chatServiceAggregator: string;
  middleTier: string;
  unifiedPresence: string;
  ams: string;
  amsV2: string;
  [key: string]: string;
}

export interface ExtractedToken {
  host: string;
  name: string;
  token: string;
  audience?: string;
  upn?: string;
  tenantId?: string;
  skypeId?: string;
  expiresAt: Date;
}

export interface MessageEmotionUser {
  mri: string;
  time?: number;
  value?: string;
  [key: string]: unknown;
}

export interface MessageEmotionEntry {
  key: string;
  users?: MessageEmotionUser[];
  [key: string]: unknown;
}

export interface MessageQuoteReference {
  messageId?: string | number;
  sender?: string;
  time?: number;
  message?: string | null;
  validationResult?: string;
  [key: string]: unknown;
}

export interface MessageActivity {
  activityType?: string;
  activityOperationType?: string;
  activitySubtype?: string;
  activityTimestamp?: string;
  activityId?: number;
  sourceThreadId?: string;
  sourceMessageId?: number;
  sourceMessageVersion?: number;
  sourceReplyChainId?: number;
  sourceUserId?: string;
  sourceUserImDisplayName?: string;
  targetUserId?: string;
  count?: string;
  messagePreview?: string;
  messagePreviewTemplateOption?: string;
  activityContext?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MessageAnnotationsSummary {
  emotions?: Record<string, number>;
  [key: string]: unknown;
}

export interface MessageProperties extends Record<string, unknown> {
  importance?: string;
  subject?: string | null;
  edittime?: string | number;
  deletetime?: number;
  hardDeleteTime?: number;
  hardDeleteReason?: string;
  composetime?: string;
  languageStamp?: string;
  formatVariant?: string;
  mentions?: string | Record<string, unknown>[];
  files?: string | Record<string, unknown>[];
  links?: string | Record<string, unknown>[];
  cards?: string | Record<string, unknown>[];
  emotions?: string | MessageEmotionEntry[];
  deltaEmotions?: string | MessageEmotionEntry[];
  qtdMsgs?: string | MessageQuoteReference[];
  pinned?: string | Record<string, unknown>;
  activity?: string | MessageActivity;
  meeting?: string | Record<string, unknown>;
  meta?: string | Record<string, unknown>;
  botMetadata?: string | Record<string, unknown>;
  botPoweredByAI?: string | boolean;
  botCitations?: string | Record<string, unknown>[];
  originalMessageContext?: string | Record<string, unknown>;
  onbehalfof?: string | Record<string, unknown>;
  atp?: string | Record<string, unknown>;
  blurHash?: string;
  counterPartyMessageId?: string | number;
  messageUpdatePolicyValue?: string | Record<string, unknown>;
  skipfanouttobots?: string | boolean;
  isread?: string | boolean;
  isSharedInMain?: string | boolean;
  issaved?: string | boolean;
  forwardTemplateId?: string;
  draftId?: string;
  announceViaEmail?: string | boolean;
  announceViaEmailPendingMembers?: string | Record<string, unknown>[];
  botFeedbackLoopEnabled?: string | boolean;
  eventReason?: string;
  "call-log"?: string | Record<string, unknown>;
}

export interface ConversationProperties extends Record<string, unknown> {
  consumptionhorizon?: string;
  consumptionHorizonBookmark?: string;
  addedBy?: string;
  addedByTenantId?: string;
  isemptyconversation?: string;
  lastimreceivedtime?: string;
  lastimportantimreceivedtime?: string;
  lasturgentimreceivedtime?: string;
  quickReplyAugmentation?: string | Record<string, unknown>;
  alerts?: string | Record<string, unknown>;
  meetingInfo?: string | Record<string, unknown>;
  draftVersion?: string | number;
  favorite?: string | boolean;
  lastTimeFavorited?: string | number;
  hasImpersonation?: string | boolean;
  collapsed?: string | boolean;
}

export interface ConversationThreadProperties extends Record<string, unknown> {
  topic?: string;
  threadType?: string;
  productThreadType?: string;
  lastjoinat?: string;
  membercount?: string;
  memberCount?: string | number;
  groupId?: string;
  spaceId?: string;
  spaceThreadTopic?: string;
  topicThreadTopic?: string;
  sharepointSiteUrl?: string;
  topics?: string;
  hidden?: string | boolean;
  rosterVersion?: string;
  originalThreadId?: string;
  createdat?: string;
  lastSequenceId?: string;
  lastleaveat?: string;
  privacy?: string;
  uniquerosterthread?: string | boolean;
  ongoingCallChatEnforcement?: string | boolean;
  tenantid?: string;
  version?: string | number;
  isCreator?: string | boolean;
  gapDetectionEnabled?: string | boolean;
}

export type TeamsSpecialThreadType =
  | "annotations"
  | "calllogs"
  | "drafts"
  | "mentions"
  | "notifications"
  | "notes";

export interface Conversation {
  id: string;
  conversationType?: string;
  type?: string;
  version?: string | number;
  properties?: ConversationProperties;
  threadProperties?: ConversationThreadProperties;
  lastMessage?: Message;
  members?: ConversationMember[];
  memberCount?: number;
  consumptionHorizon?: string;
  lastActivityTime?: string;
  specialThreadType?: TeamsSpecialThreadType;
}

export interface ConversationMember {
  id: string;
  role: string;
  isMri: boolean;
  displayName?: string;
  friendlyName?: string;
  tenantId?: string;
  objectId?: string;
  userPrincipalName?: string;
  [key: string]: unknown;
}

export interface MemberConsumptionHorizon {
  id: string;
  consumptionhorizon: string;
  messageVisibilityTime?: number;
  [key: string]: unknown;
}

export interface MembersConsumptionHorizonsResponse {
  id: string;
  version?: string;
  consumptionhorizons: MemberConsumptionHorizon[];
  [key: string]: unknown;
}

export interface Message {
  id: string;
  sequenceId?: number;
  clientMessageId?: string;
  version?: string;
  conversationId: string;
  type: string;
  messagetype: string;
  contenttype: string;
  content?: string;
  from: string;
  fromMri?: string | null;
  fromTenantId?: string;
  imdisplayname?: string;
  senderDisplayName?: string;
  composetime: string;
  originalarrivaltime: string;
  timestamp?: string;
  deleted?: boolean;
  annotationsSummary?: MessageAnnotationsSummary;
  amsreferences?: unknown[];
  conversationLink?: string;
  postType?: string;
  s2spartnername?: string;
  skypeguid?: string;
  rootMessageId?: string;
  clumpId?: string;
  secondaryReferenceId?: string;
  prioritizeImDisplayName?: boolean;
  fromDisplayNameInToken?: string;
  fromGivenNameInToken?: string;
  fromFamilyNameInToken?: string;
  properties?: MessageProperties;
}

export interface MessagesResponse {
  messages: Message[];
  _metadata?: {
    backwardLink?: string;
    syncState?: string;
    lastCompleteSegmentStartTime?: number;
    lastCompleteSegmentEndTime?: number;
  };
}

export interface ConversationsResponse {
  conversations: Conversation[];
  _metadata?: {
    backwardLink?: string;
    syncState?: string;
  };
}

export type TeamsProfilePresentation = {
  avatarThumbs: Record<string, string>;
  avatarFull: Record<string, string>;
  displayNames: Record<string, string>;
  emails: Record<string, string>;
  jobTitles: Record<string, string>;
  departments: Record<string, string>;
  companyNames: Record<string, string>;
  tenantNames: Record<string, string>;
  locations: Record<string, string>;
};

export interface PresenceInfo {
  availability?: string;
  activity?: string;
  statusMessage?: {
    message: string;
    expiry: string;
  };
}
