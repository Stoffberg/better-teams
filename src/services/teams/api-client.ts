/**
 * Client for the Microsoft Teams internal APIs.
 *
 * Uses tokens extracted from the local Teams installation to call
 * the same undocumented APIs that the Teams desktop client uses.
 *
 * Auth flow:
 *   1. Extract Bearer token (aud: api.spaces.skype.com) from Teams cookies
 *   2. POST to /api/authsvc/v1.0/authz → get skypeToken + regionGtms
 *   3. Use skypeToken for chat service APIs (ng.msg.teams)
 *   4. Use Bearer token for CSA, MT, and presence APIs
 */

import {
  filterConversationsForPipeline,
  sortConversationsByActivity,
} from "@/lib/chat-format";
import {
  cacheImageFile,
  extractTokens,
  filePathToAssetUrl,
  getAuthToken,
} from "@/lib/tauri-bridge";
import {
  applyProfileDisplayNameToRowMrIs,
  applyProfilePhotoDataUrlToRowMrIs,
  canonAvatarMri,
  companyNameFromShortProfileRow,
  departmentFromShortProfileRow,
  displayNameFromShortProfileRow,
  emailFromShortProfileRow,
  jobTitleFromShortProfileRow,
  locationFromShortProfileRow,
  mriFromShortProfileRow,
  shortProfileRowToMriAndImageUrl,
  tenantNameFromShortProfileRow,
} from "@/lib/teams-profile-avatars";
import {
  AuthzResponseSchema,
  ConversationSchema,
  ConversationsResponseSchema,
  MembersConsumptionHorizonsResponseSchema,
  MessageSchema,
  MessagesResponseSchema,
  PresenceEntrySchema,
  PresenceListSchema,
  parseShortProfileRows,
  type ShortProfileRow,
  UserProfileResponseSchema,
} from "./schemas";
import type {
  AuthzResponse,
  Conversation,
  ConversationMember,
  ConversationsResponse,
  ExtractedToken,
  MembersConsumptionHorizonsResponse,
  Message,
  MessagesResponse,
  PresenceInfo,
  RegionGtms,
  TeamsProfilePresentation,
} from "./types";

const AUTHZ_URL = "https://teams.microsoft.com/api/authsvc/v1.0/authz";

const TEAMS_WEB_UA =
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36";

const DEFAULT_ASM_URL = "https://api.asm.skype.com";

type AttachmentKind = "image" | "file";

type AmsObjectCreateResponse = {
  id: string;
};

export type TeamsApiDiagnosticEvent = {
  level: "warn" | "error";
  message: string;
  data?: Record<string, unknown>;
};

export type TeamsApiClientOptions = {
  fetchImpl?: (
    input: RequestInfo | URL,
    init?: RequestInit,
  ) => Promise<Response>;
  onDiagnostic?: (event: TeamsApiDiagnosticEvent) => void;
  getCachedImagePath?: (url: string) => string | Promise<string | null> | null;
  setCachedImagePath?: (url: string, filePath: string) => void | Promise<void>;
};

function imageFileExtensionFromMime(mime: string): string | null {
  if (mime === "image/jpeg") return "jpg";
  if (mime === "image/png") return "png";
  if (mime === "image/gif") return "gif";
  if (mime === "image/webp") return "webp";
  if (mime === "image/avif") return "avif";
  return null;
}

function sniffImageMimeFromBuffer(buf: ArrayBuffer): string | null {
  const u = new Uint8Array(buf);
  if (u.length >= 3 && u[0] === 0xff && u[1] === 0xd8 && u[2] === 0xff) {
    return "image/jpeg";
  }
  if (
    u.length >= 8 &&
    u[0] === 0x89 &&
    u[1] === 0x50 &&
    u[2] === 0x4e &&
    u[3] === 0x47
  ) {
    return "image/png";
  }
  if (u.length >= 6 && u[0] === 0x47 && u[1] === 0x49 && u[2] === 0x46) {
    return "image/gif";
  }
  if (
    u.length >= 12 &&
    u[0] === 0x52 &&
    u[1] === 0x49 &&
    u[2] === 0x46 &&
    u[8] === 0x57 &&
    u[9] === 0x45 &&
    u[10] === 0x42 &&
    u[11] === 0x50
  ) {
    return "image/webp";
  }
  return null;
}

function decodeJwtPayload(token: string): Record<string, unknown> | null {
  try {
    const parts = token.split(".");
    if (parts.length < 2) return null;
    // Browser-native base64url decode (replaces Node Buffer)
    const b64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const payload = new TextDecoder().decode(
      Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)),
    );
    return JSON.parse(payload) as Record<string, unknown>;
  } catch {
    return null;
  }
}

function escapeXmlText(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value)
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function attachmentKindFromFile(file: File): AttachmentKind {
  return file.type.startsWith("image/") ? "image" : "file";
}

function normalizedAttachmentMimeType(file: File): string {
  const trimmed = file.type.trim();
  if (trimmed) return trimmed;
  return attachmentKindFromFile(file) === "image"
    ? "image/jpeg"
    : "application/octet-stream";
}

function normalizeParticipantIds(ids: string[]): string[] {
  const seen = new Set<string>();
  const normalized: string[] = [];
  for (const id of ids) {
    const trimmed = id.trim();
    if (!trimmed.startsWith("8:")) continue;
    if (seen.has(trimmed)) continue;
    seen.add(trimmed);
    normalized.push(trimmed);
  }
  return normalized;
}

function buildAttachmentMarkup(
  asmBaseUrl: string,
  objectUrl: string,
  objectId: string,
  file: File,
  kind: AttachmentKind,
): {
  content: string;
  messageType: "RichText/UriObject" | "RichText/Media_GenericFile";
} {
  if (kind === "image") {
    const viewLink = `${asmBaseUrl}/s/i?${encodeURIComponent(objectId)}`;
    return {
      messageType: "RichText/UriObject",
      content: `<URIObject type="Picture.1" uri="${escapeXmlAttribute(objectUrl)}" url_thumbnail="${escapeXmlAttribute(`${objectUrl}/views/imgt1`)}"><Title/><Description/><OriginalName v="${escapeXmlAttribute(file.name)}"/><a href="${escapeXmlAttribute(viewLink)}">${escapeXmlText(viewLink)}</a><meta type="photo" originalName="${escapeXmlAttribute(file.name)}"/></URIObject>`,
    };
  }

  const viewLink = `https://login.skype.com/login/sso?go=webclient.xmm&docid=${encodeURIComponent(objectId)}`;
  return {
    messageType: "RichText/Media_GenericFile",
    content: `<URIObject type="File.1" uri="${escapeXmlAttribute(objectUrl)}" url_thumbnail="${escapeXmlAttribute(`${objectUrl}/views/thumbnail`)}"><Title>${escapeXmlText(`Title: ${file.name}`)}</Title><Description>${escapeXmlText(`Description: ${file.name}`)}</Description><FileSize v="${file.size}"/><OriginalName v="${escapeXmlAttribute(file.name)}"/><a href="${escapeXmlAttribute(viewLink)}">${escapeXmlText(viewLink)}</a></URIObject>`,
  };
}

export class TeamsApiClient {
  private authToken: ExtractedToken | null = null;
  private skypeToken: string | null = null;
  private regionGtms: RegionGtms | null = null;
  private region: string | null = null;
  private tenantId: string | undefined;
  private readonly httpFetch: (
    input: RequestInfo | URL,
    init?: RequestInit,
  ) => Promise<Response>;
  private readonly onDiagnostic?: (event: TeamsApiDiagnosticEvent) => void;
  private readonly getCachedImagePath?: (
    url: string,
  ) => string | Promise<string | null> | null;
  private readonly setCachedImagePath?: (
    url: string,
    filePath: string,
  ) => void | Promise<void>;
  private presenceAvailable = true;

  constructor(tenantId?: string, options?: TeamsApiClientOptions) {
    this.tenantId = tenantId;
    this.httpFetch =
      options?.fetchImpl ??
      ((input, init) => globalThis.fetch(input as RequestInfo, init));
    this.onDiagnostic = options?.onDiagnostic;
    this.getCachedImagePath = options?.getCachedImagePath;
    this.setCachedImagePath = options?.setCachedImagePath;
  }

  private emitDiagnostic(event: TeamsApiDiagnosticEvent): void {
    this.onDiagnostic?.(event);
  }

  private companionHeadersForTeamsMicrosoftCom(
    url: string,
  ): Record<string, string> {
    try {
      const u = new URL(url);
      const host = u.hostname;
      if (host.includes("msg.teams.microsoft.com")) {
        return {};
      }
      const middleTierPath =
        host === "teams.microsoft.com" ||
        (/\.teams\.microsoft\.com$/i.test(host) &&
          u.pathname.includes("/api/mt/"));
      if (!middleTierPath && host !== "teams.microsoft.com") {
        return {};
      }
      return {
        "User-Agent": TEAMS_WEB_UA,
        Origin: "https://teams.microsoft.com",
        Referer: "https://teams.microsoft.com/",
      };
    } catch {
      return {};
    }
  }

  /**
   * Initialize the client by extracting tokens and calling the authz endpoint.
   * Must be called before any API methods.
   */
  async initialize(): Promise<void> {
    this.authToken = await getAuthToken(this.tenantId);
    if (!this.authToken) {
      throw new Error(
        "No valid Teams auth token found. Ensure Teams is running and signed in.",
      );
    }

    const authz = await this.callAuthz();
    this.skypeToken = authz.tokens.skypeToken;
    this.regionGtms = authz.regionGtms;
    this.region = authz.region;
    if (!this.authToken.skypeId && this.skypeToken) {
      const payload = decodeJwtPayload(this.skypeToken);
      const fromSkypeToken =
        payload && typeof payload.skypeid === "string" ? payload.skypeid : null;
      if (fromSkypeToken) {
        this.authToken.skypeId = fromSkypeToken;
      }
    }
  }

  /**
   * Check if the current tokens are still valid.
   */
  get isAuthenticated(): boolean {
    if (!this.authToken || !this.skypeToken) return false;
    return this.authToken.expiresAt > new Date();
  }

  /**
   * Re-extract tokens from Teams cookie store if current ones are expired.
   */
  async refreshIfNeeded(): Promise<void> {
    if (!this.isAuthenticated) {
      await this.initialize();
    }
  }

  /**
   * Get info about the currently authenticated account.
   */
  get account() {
    return {
      upn: this.authToken?.upn,
      tenantId: this.authToken?.tenantId,
      expiresAt: this.authToken?.expiresAt,
      region: this.region,
      skypeId: this.authToken?.skypeId,
    };
  }

  /**
   * Get all available accounts (across tenants).
   */
  static async getAvailableAccounts(): Promise<
    { upn?: string; tenantId?: string }[]
  > {
    const tokens = await extractTokens();
    const authTokens = tokens.filter((t) => t.name === "authtoken");
    const seen = new Set<string>();
    const accounts: { upn?: string; tenantId?: string }[] = [];

    for (const t of authTokens) {
      const key = `${t.tenantId}:${t.upn}`;
      if (!seen.has(key)) {
        seen.add(key);
        accounts.push({ upn: t.upn, tenantId: t.tenantId });
      }
    }

    return accounts;
  }

  // ── Auth ──

  private async callAuthz(): Promise<AuthzResponse> {
    const res = await this.httpFetch(AUTHZ_URL, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${this.authToken?.token}`,
        "ms-teams-authz-type": "TokenRefresh",
        ...this.companionHeadersForTeamsMicrosoftCom(AUTHZ_URL),
      },
    });

    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`Authz failed (${res.status}): ${body.slice(0, 200)}`);
    }

    return AuthzResponseSchema.parse(await res.json());
  }

  // ── Conversations ──

  /**
   * Fetch recent conversations (1:1 chats, group chats) via the chat service.
   */
  async getConversations(pageSize = 20): Promise<ConversationsResponse> {
    const parsed = await this.getConversationsPageUnfiltered(
      pageSize,
      "0",
      "msnp24Equivalent",
    );
    const list = parsed.conversations ?? [];
    return {
      ...parsed,
      conversations: filterConversationsForPipeline(list),
    };
  }

  async getConversationsPageUnfiltered(
    pageSize = 40,
    startTime: string | number = 0,
    view = "msnp24Equivalent",
  ): Promise<ConversationsResponse> {
    await this.refreshIfNeeded();
    const url = `${this.regionGtms?.chatService}/v1/users/ME/conversations?view=${encodeURIComponent(view)}&pageSize=${pageSize}&startTime=${startTime}`;
    return ConversationsResponseSchema.parse(
      await this.fetchWithSkypeToken(url),
    );
  }

  async getConversationsByUrl(url: string): Promise<ConversationsResponse> {
    await this.refreshIfNeeded();
    return ConversationsResponseSchema.parse(
      await this.fetchWithSkypeToken(url),
    );
  }

  async getAllConversations(pageSize = 100): Promise<ConversationsResponse> {
    return this.getAllConversationsForView("msnp24Equivalent", pageSize);
  }

  async getFavoriteConversations(
    pageSize = 100,
  ): Promise<ConversationsResponse> {
    return this.getAllConversationsForView("favorites", pageSize);
  }

  async setConversationFavorite(
    conversationId: string,
    favorite: boolean,
  ): Promise<void> {
    await this.refreshIfNeeded();
    const encodedConversationId = encodeURIComponent(conversationId);
    const url = `${this.regionGtms?.chatService}/v1/users/ME/conversations/${encodedConversationId}/properties`;
    await this.sendWithSkypeToken(url, {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ favorite }),
    });
  }

  private async getAllConversationsForView(
    view: string,
    pageSize = 100,
  ): Promise<ConversationsResponse> {
    const seenConversationIds = new Set<string>();
    const conversations: Conversation[] = [];
    let cursor: string | number = 0;
    let lastResponse: ConversationsResponse | null = null;

    for (let page = 0; page < 20; page += 1) {
      const response: ConversationsResponse =
        typeof cursor === "string" && /^https?:\/\//i.test(cursor)
          ? await this.getConversationsByUrl(cursor)
          : await this.getConversationsPageUnfiltered(pageSize, cursor, view);
      lastResponse = response;
      const pageConversations = response.conversations ?? [];
      for (const conversation of pageConversations) {
        if (seenConversationIds.has(conversation.id)) continue;
        seenConversationIds.add(conversation.id);
        conversations.push(conversation);
      }
      const nextCursor: string | undefined =
        response._metadata?.backwardLink?.trim() ??
        response._metadata?.syncState?.trim();
      if (!nextCursor || nextCursor === String(cursor)) break;
      cursor = nextCursor;
      if (pageConversations.length === 0) break;
    }

    return {
      ...(lastResponse ?? { conversations: [] }),
      conversations: filterConversationsForPipeline(
        sortConversationsByActivity(conversations),
      ),
    };
  }

  async getConversation(conversationId: string): Promise<Conversation> {
    await this.refreshIfNeeded();
    const encodedId = encodeURIComponent(conversationId);
    const url = `${this.regionGtms?.chatService}/v1/users/ME/conversations/${encodedId}`;
    return ConversationSchema.parse(await this.fetchWithSkypeToken(url));
  }

  async getThreadMembers(
    conversationId: string,
  ): Promise<ConversationMember[]> {
    await this.refreshIfNeeded();
    const encodedId = encodeURIComponent(conversationId);
    const url = `${this.regionGtms?.chatService}/v1/threads/${encodedId}/members`;
    const response = (await this.fetchWithSkypeToken(url)) as {
      members?: Array<Record<string, unknown>>;
    };
    return (response.members ?? [])
      .map((member) => {
        const id = typeof member.id === "string" ? member.id.trim() : "";
        if (!id) return null;
        return {
          id,
          role: typeof member.role === "string" ? member.role : "User",
          isMri: id.startsWith("8:"),
          ...(typeof member.displayName === "string"
            ? { displayName: member.displayName }
            : {}),
          ...(typeof member.friendlyName === "string"
            ? { friendlyName: member.friendlyName }
            : {}),
          ...(typeof member.userPrincipalName === "string"
            ? { userPrincipalName: member.userPrincipalName }
            : {}),
        } satisfies ConversationMember;
      })
      .filter((member): member is ConversationMember => member != null);
  }

  async getMembersConsumptionHorizon(
    conversationId: string,
  ): Promise<MembersConsumptionHorizonsResponse> {
    await this.refreshIfNeeded();
    const encodedId = encodeURIComponent(conversationId);
    const url = `${this.regionGtms?.chatService}/v1/threads/${encodedId}/consumptionhorizons`;
    return MembersConsumptionHorizonsResponseSchema.parse(
      await this.fetchWithSkypeToken(url),
    );
  }

  // ── Messages ──

  /**
   * Fetch messages from a specific conversation.
   */
  async getMessages(
    conversationId: string,
    pageSize = 50,
    startTime = 1,
  ): Promise<MessagesResponse> {
    await this.refreshIfNeeded();
    const encodedId = encodeURIComponent(conversationId);
    const url = `${this.regionGtms?.chatService}/v1/users/ME/conversations/${encodedId}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=${pageSize}&startTime=${startTime}`;
    return MessagesResponseSchema.parse(await this.fetchWithSkypeToken(url));
  }

  async getMessagesByUrl(url: string): Promise<MessagesResponse> {
    await this.refreshIfNeeded();
    return MessagesResponseSchema.parse(await this.fetchWithSkypeToken(url));
  }

  async getMessagesAroundMessage(
    conversationId: string,
    messageId: string,
  ): Promise<Conversation> {
    await this.refreshIfNeeded();
    const encodedConversationId = encodeURIComponent(conversationId);
    const encodedMessageId = encodeURIComponent(messageId);
    const url = `${this.regionGtms?.chatService}/v1/users/ME/conversations/${encodedConversationId};messageid=${encodedMessageId}`;
    return ConversationSchema.parse(await this.fetchWithSkypeToken(url));
  }

  async getAnchoredMessages(
    conversationId: string,
    messageId: string,
  ): Promise<MessagesResponse> {
    await this.refreshIfNeeded();
    const encodedConversationId = encodeURIComponent(conversationId);
    const encodedMessageId = encodeURIComponent(messageId);
    const url = `${this.regionGtms?.chatService}/v1/users/ME/conversations/${encodedConversationId};messageid=${encodedMessageId}/messages`;
    return MessagesResponseSchema.parse(await this.fetchWithSkypeToken(url));
  }

  async getMessage(
    conversationId: string,
    messageId: string,
  ): Promise<Message> {
    await this.refreshIfNeeded();
    const encodedConversationId = encodeURIComponent(conversationId);
    const encodedMessageId = encodeURIComponent(messageId);
    const url = `${this.regionGtms?.chatService}/v1/users/ME/conversations/${encodedConversationId}/messages/${encodedMessageId}`;
    return MessageSchema.parse(await this.fetchWithSkypeToken(url));
  }

  /**
   * Send a text message to a conversation.
   */
  async sendMessage(
    conversationId: string,
    content: string,
    displayName: string,
    contentFormat: "html" | "text" = "text",
    mentions: Array<Record<string, unknown>> = [],
  ): Promise<void> {
    await this.refreshIfNeeded();
    const encodedId = encodeURIComponent(conversationId);
    const url = `${this.regionGtms?.chatService}/v1/users/ME/conversations/${encodedId}/messages`;

    const body = {
      content,
      messagetype: contentFormat === "html" ? "RichText/Html" : "Text",
      contenttype: contentFormat === "html" ? "RichText/Html" : "text",
      amsreferences: [],
      clientmessageid: Date.now().toString(),
      imdisplayname: displayName,
      properties: {
        importance: "",
        subject: null,
        ...(mentions.length > 0 ? { mentions } : {}),
        ...(contentFormat === "html" ? { formatVariant: "RichText/Html" } : {}),
      },
    };

    const res = await this.httpFetch(url, {
      method: "POST",
      headers: {
        Authentication: `skypetoken=${this.skypeToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    if (!res.ok) {
      const err = await res.text().catch(() => "");
      throw new Error(
        `Send message failed (${res.status}): ${err.slice(0, 200)}`,
      );
    }
  }

  async sendAttachmentMessage(
    conversationId: string,
    file: File,
    displayName: string,
    participantIds?: string[],
  ): Promise<void> {
    await this.refreshIfNeeded();
    const kind = attachmentKindFromFile(file);
    const attachmentParticipants = await this.resolveAttachmentParticipantIds(
      conversationId,
      participantIds,
    );
    const objectUrl = this.regionGtms?.ams ?? DEFAULT_ASM_URL;
    const createRes = await this.httpFetch(`${objectUrl}/v1/objects`, {
      method: "POST",
      headers: {
        Authorization: `skype_token ${this.skypeToken}`,
        "Content-Type": "application/json",
        "X-Client-Version": "0/0.0.0.0",
      },
      body: JSON.stringify({
        type: kind === "image" ? "pish/image" : "sharing/file",
        permissions: Object.fromEntries(
          attachmentParticipants.map((id) => [id, ["read"]]),
        ),
        ...(kind === "file" ? { filename: file.name } : {}),
      }),
    });

    if (!createRes.ok) {
      const err = await createRes.text().catch(() => "");
      throw new Error(
        `Attachment create failed (${createRes.status}): ${err.slice(0, 200)}`,
      );
    }

    const created = (await createRes.json()) as AmsObjectCreateResponse;
    const objectId = created.id?.trim();
    if (!objectId) {
      throw new Error("Attachment create failed: missing object id");
    }

    const attachmentUrl = `${objectUrl}/v1/objects/${encodeURIComponent(objectId)}`;
    const uploadRes = await this.httpFetch(
      `${attachmentUrl}/content/${kind === "image" ? "imgpsh" : "original"}`,
      {
        method: "PUT",
        headers: {
          Authorization: `skype_token ${this.skypeToken}`,
          "Content-Type": normalizedAttachmentMimeType(file),
        },
        body: await file.arrayBuffer(),
      },
    );

    if (!uploadRes.ok) {
      const err = await uploadRes.text().catch(() => "");
      throw new Error(
        `Attachment upload failed (${uploadRes.status}): ${err.slice(0, 200)}`,
      );
    }

    const message = buildAttachmentMarkup(
      objectUrl,
      attachmentUrl,
      objectId,
      file,
      kind,
    );
    const encodedId = encodeURIComponent(conversationId);
    const url = `${this.regionGtms?.chatService}/v1/users/ME/conversations/${encodedId}/messages`;
    const res = await this.httpFetch(url, {
      method: "POST",
      headers: {
        Authentication: `skypetoken=${this.skypeToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        content: message.content,
        messagetype: message.messageType,
        contenttype: message.messageType,
        amsreferences: [objectId],
        clientmessageid: Date.now().toString(),
        imdisplayname: displayName,
        properties: {
          formatVariant: message.messageType,
        },
      }),
    });

    if (!res.ok) {
      const err = await res.text().catch(() => "");
      throw new Error(
        `Send attachment failed (${res.status}): ${err.slice(0, 200)}`,
      );
    }
  }

  // ── Message actions ──

  /**
   * Delete (soft-delete) a message from a conversation.
   */
  async deleteMessage(
    conversationId: string,
    messageId: string,
  ): Promise<void> {
    await this.refreshIfNeeded();
    const encodedConversationId = encodeURIComponent(conversationId);
    const encodedMessageId = encodeURIComponent(messageId);
    const url = `${this.regionGtms?.chatService}/v1/users/ME/conversations/${encodedConversationId}/messages/${encodedMessageId}`;

    const res = await this.httpFetch(url, {
      method: "DELETE",
      headers: {
        Authentication: `skypetoken=${this.skypeToken}`,
      },
    });

    if (!res.ok) {
      const err = await res.text().catch(() => "");
      throw new Error(
        `Delete message failed (${res.status}): ${err.slice(0, 200)}`,
      );
    }
  }

  // ── Presence ──

  /**
   * Get presence (availability) for one or more users by their MRI.
   * MRI format: "8:orgid:{user-guid}"
   */
  async getPresence(userMris: string[]): Promise<Record<string, PresenceInfo>> {
    if (!this.presenceAvailable || userMris.length === 0) return {};
    await this.refreshIfNeeded();
    const url = `${this.regionGtms?.unifiedPresence}/v1/presence/getpresence`;

    const res = await this.httpFetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${this.authToken?.token}`,
        "Content-Type": "application/json",
        ...this.companionHeadersForTeamsMicrosoftCom(url),
      },
      body: JSON.stringify(userMris.map((mri) => ({ mri }))),
    });

    if (!res.ok) {
      if (res.status === 401) {
        this.presenceAvailable = false;
        this.emitDiagnostic({
          level: "warn",
          message: "Presence API unavailable for current tenant",
          data: { status: res.status, tenantId: this.tenantId },
        });
        return {};
      }
      throw new Error(`Presence request failed (${res.status})`);
    }

    const raw = await res.json();
    const data = PresenceListSchema.safeParse(raw);
    const result: Record<string, PresenceInfo> = {};

    const entries = data.success
      ? data.data
      : Array.isArray(raw)
        ? raw.flatMap((entry) => {
            const parsedEntry = PresenceEntrySchema.safeParse(entry);
            return parsedEntry.success ? [parsedEntry.data] : [];
          })
        : [];

    const byMri = new Map(
      entries
        .filter(
          (entry) =>
            typeof entry.mri === "string" &&
            entry.mri.trim() &&
            (entry.presence.availability || entry.presence.activity),
        )
        .map((entry) => [entry.mri.trim().toLowerCase(), entry.presence]),
    );
    for (const userMri of userMris) {
      const presence = byMri.get(userMri.trim().toLowerCase());
      if (!presence) continue;
      result[userMri] = presence;
    }
    return result;
  }

  /**
   * Set your own availability.
   */
  async setAvailability(
    availability:
      | "Available"
      | "Busy"
      | "DoNotDisturb"
      | "BeRightBack"
      | "Away"
      | "Offline",
  ): Promise<void> {
    if (!this.presenceAvailable) return;
    await this.refreshIfNeeded();
    const url = `${this.regionGtms?.unifiedPresence}/v1/me/forceavailability/`;

    const res = await this.httpFetch(url, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${this.authToken?.token}`,
        "Content-Type": "application/json",
        ...this.companionHeadersForTeamsMicrosoftCom(url),
      },
      body: JSON.stringify({ availability }),
    });

    if (!res.ok) {
      if (res.status === 401) {
        this.presenceAvailable = false;
        this.emitDiagnostic({
          level: "warn",
          message: "Self availability API unavailable for current tenant",
          data: { status: res.status, tenantId: this.tenantId },
        });
        return;
      }
      throw new Error(`Set availability failed (${res.status})`);
    }
  }

  // ── User Profiles ──

  /**
   * Get user profile by email via the Middle Tier API.
   */
  async getUserProfile(email: string): Promise<unknown> {
    await this.refreshIfNeeded();
    const url = `${this.regionGtms?.middleTier}/beta/users/${encodeURIComponent(email)}/?throwIfNotFound=false&isMailAddress=true&enableGuest=true&includeIBBarredUsers=true&skypeTeamsInfo=true`;
    return UserProfileResponseSchema.parse(await this.fetchWithBearer(url));
  }

  /**
   * Batch fetch short profiles by MRI.
   */
  async fetchShortProfiles(mris: string[]): Promise<ShortProfileRow[]> {
    await this.refreshIfNeeded();
    const url = `${this.regionGtms?.middleTier}/beta/users/fetchShortProfile?isMailAddress=false&enableGuest=true&includeIBBarredUsers=false&skypeTeamsInfo=true`;

    const res = await this.httpFetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${this.authToken?.token}`,
        "Content-Type": "application/json",
        Accept: "application/json",
        ...this.companionHeadersForTeamsMicrosoftCom(url),
      },
      body: JSON.stringify(mris),
    });

    if (!res.ok) {
      throw new Error(`Fetch profiles failed (${res.status})`);
    }

    return parseShortProfileRows(await res.json());
  }

  async fetchProfileAvatarDataUrls(
    mris: string[],
  ): Promise<TeamsProfilePresentation> {
    const unique = [...new Set(mris.map((m) => m.trim()).filter(Boolean))];
    if (unique.length === 0) {
      return {
        avatarThumbs: {},
        avatarFull: {},
        displayNames: {},
        emails: {},
        jobTitles: {},
        departments: {},
        companyNames: {},
        tenantNames: {},
        locations: {},
      };
    }

    await this.refreshIfNeeded();
    const avatarThumbs: Record<string, string> = {};
    const avatarFull: Record<string, string> = {};
    const displayNames: Record<string, string> = {};
    const emails: Record<string, string> = {};
    const jobTitles: Record<string, string> = {};
    const departments: Record<string, string> = {};
    const companyNames: Record<string, string> = {};
    const tenantNames: Record<string, string> = {};
    const locations: Record<string, string> = {};
    const batchSize = 40;

    const setAvatarThumb = (mri: string, dataUrl: string) => {
      avatarThumbs[canonAvatarMri(mri)] = dataUrl;
    };
    const setAvatarFull = (mri: string, dataUrl: string) => {
      avatarFull[canonAvatarMri(mri)] = dataUrl;
    };
    const setDisplayName = (mri: string, label: string) => {
      displayNames[canonAvatarMri(mri)] = label;
    };
    const hasAvatarThumb = (mri: string) =>
      Object.hasOwn(avatarThumbs, canonAvatarMri(mri));
    const hasAvatarFull = (mri: string) =>
      Object.hasOwn(avatarFull, canonAvatarMri(mri));

    const imageFetchConcurrency = 12;

    for (let i = 0; i < unique.length; i += batchSize) {
      const chunk = unique.slice(i, i + batchSize);
      let rows: ShortProfileRow[];
      try {
        rows = await this.fetchShortProfiles(chunk);
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        this.emitDiagnostic({
          level: "warn",
          message: "teams.fetchShortProfile batch failed",
          data: {
            error: msg,
            batchStart: i,
            batchSize: chunk.length,
          },
        });
        rows = [];
      }
      for (const row of rows) {
        const dn = displayNameFromShortProfileRow(row);
        if (dn) {
          applyProfileDisplayNameToRowMrIs(row, dn, setDisplayName);
        }
        const rowMriForMeta = mriFromShortProfileRow(row);
        if (rowMriForMeta) {
          const em = emailFromShortProfileRow(row);
          if (em) emails[canonAvatarMri(rowMriForMeta)] = em;
          const jt = jobTitleFromShortProfileRow(row);
          if (jt) jobTitles[canonAvatarMri(rowMriForMeta)] = jt;
          const dept = departmentFromShortProfileRow(row);
          if (dept) departments[canonAvatarMri(rowMriForMeta)] = dept;
          const company = companyNameFromShortProfileRow(row);
          if (company) companyNames[canonAvatarMri(rowMriForMeta)] = company;
          const tenant = tenantNameFromShortProfileRow(row);
          if (tenant) tenantNames[canonAvatarMri(rowMriForMeta)] = tenant;
          const location = locationFromShortProfileRow(row);
          if (location) locations[canonAvatarMri(rowMriForMeta)] = location;
        }
      }
      const jobs: Array<{ row: unknown; url: string }> = [];
      const rowPicFallbacks: Array<{ row: unknown; rowMri: string }> = [];
      for (const row of rows) {
        const parsed = shortProfileRowToMriAndImageUrl(row);
        if (parsed) {
          jobs.push({ row, url: parsed.imageUrl });
          continue;
        }
        const rowMri = mriFromShortProfileRow(row);
        if (rowMri) rowPicFallbacks.push({ row, rowMri });
      }
      await Promise.all(
        jobs.map(async ({ row, url }) => {
          const imageSrc = await this.fetchAuthenticatedImageSrc(url);
          if (imageSrc) {
            applyProfilePhotoDataUrlToRowMrIs(row, imageSrc, setAvatarThumb);
            applyProfilePhotoDataUrlToRowMrIs(row, imageSrc, setAvatarFull);
          }
        }),
      );

      for (let j = 0; j < rowPicFallbacks.length; j += imageFetchConcurrency) {
        const slice = rowPicFallbacks.slice(j, j + imageFetchConcurrency);
        await Promise.all(
          slice.map(async ({ row, rowMri }) => {
            const [thumbDataUrl, fullDataUrl] = await Promise.all([
              this.fetchFirstProfilePictureDataUrl(rowMri, "thumb"),
              this.fetchFirstProfilePictureDataUrl(rowMri, "full"),
            ]);
            if (thumbDataUrl || fullDataUrl) {
              applyProfilePhotoDataUrlToRowMrIs(
                row,
                thumbDataUrl ?? fullDataUrl ?? "",
                setAvatarThumb,
              );
              applyProfilePhotoDataUrlToRowMrIs(
                row,
                fullDataUrl ?? thumbDataUrl ?? "",
                setAvatarFull,
              );
            }
          }),
        );
      }

      const stillMissing = chunk.filter(
        (mri) => !hasAvatarThumb(mri) || !hasAvatarFull(mri),
      );
      for (let j = 0; j < stillMissing.length; j += imageFetchConcurrency) {
        const slice = stillMissing.slice(j, j + imageFetchConcurrency);
        await Promise.all(
          slice.map(async (mri) => {
            const [thumbDataUrl, fullDataUrl] = await Promise.all([
              hasAvatarThumb(mri)
                ? Promise.resolve(avatarThumbs[canonAvatarMri(mri)])
                : this.fetchFirstProfilePictureDataUrl(mri, "thumb"),
              hasAvatarFull(mri)
                ? Promise.resolve(avatarFull[canonAvatarMri(mri)])
                : this.fetchFirstProfilePictureDataUrl(mri, "full"),
            ]);
            if (thumbDataUrl) setAvatarThumb(mri, thumbDataUrl);
            if (fullDataUrl) setAvatarFull(mri, fullDataUrl);
            if (!thumbDataUrl && fullDataUrl) setAvatarThumb(mri, fullDataUrl);
            if (!fullDataUrl && thumbDataUrl) setAvatarFull(mri, thumbDataUrl);
          }),
        );
      }

      const rowsMissingBoth = rows.filter((row) => {
        const rowMri = mriFromShortProfileRow(row);
        return rowMri
          ? !hasAvatarThumb(rowMri) || !hasAvatarFull(rowMri)
          : false;
      });
      if (rowsMissingBoth.length > 0) {
        await Promise.all(
          rowsMissingBoth.map(async (row) => {
            const rowMri = mriFromShortProfileRow(row);
            if (!rowMri) return;
            const fallback =
              (hasAvatarFull(rowMri) && avatarFull[canonAvatarMri(rowMri)]) ||
              (hasAvatarThumb(rowMri) && avatarThumbs[canonAvatarMri(rowMri)]);
            if (!fallback) return;
            if (!hasAvatarThumb(rowMri)) {
              applyProfilePhotoDataUrlToRowMrIs(row, fallback, setAvatarThumb);
            }
            if (!hasAvatarFull(rowMri)) {
              applyProfilePhotoDataUrlToRowMrIs(row, fallback, setAvatarFull);
            }
          }),
        );
      }
    }

    return {
      avatarThumbs,
      avatarFull,
      displayNames,
      emails,
      jobTitles,
      departments,
      companyNames,
      tenantNames,
      locations,
    };
  }

  // ── Internal helpers ──

  private profilePictureUrlCandidates(
    mri: string,
    quality: "thumb" | "full",
  ): string[] {
    const mt = this.regionGtms?.middleTier;
    if (!mt) return [];
    const baseMt = mt.replace(/\/$/, "");
    const enc = encodeURIComponent(mri);
    const pic = `${baseMt}/v1/users/${enc}/profilePicture`;
    const q = "displayName=TeamsUser";
    return quality === "thumb"
      ? [`${pic}?${q}&size=HR64x64`, `${pic}?${q}&size=HR96x96`]
      : [`${pic}?${q}&size=HR360x360`, `${pic}?${q}&size=HR96x96`];
  }

  private async fetchFirstProfilePictureDataUrl(
    mri: string,
    quality: "thumb" | "full",
  ): Promise<string | null> {
    for (const picUrl of this.profilePictureUrlCandidates(mri, quality)) {
      const imageSrc = await this.fetchAuthenticatedImageSrc(picUrl);
      if (imageSrc) return imageSrc;
    }
    return null;
  }

  private preferSkypeTokenFirstForImageFetch(url: string): boolean {
    try {
      const h = new URL(url).hostname.toLowerCase();
      if (h.includes("asm.skype")) return true;
      if (h.includes("skype")) return true;
      if (h.includes("azureedge")) return true;
      if (h.includes("microsoftusercontent")) return true;
      if (h.includes("office.net")) return true;
      return false;
    } catch {
      return false;
    }
  }

  private usesAsmImageAuth(url: string): boolean {
    try {
      return new URL(url).hostname.toLowerCase().includes("asm.skype");
    } catch {
      return false;
    }
  }

  private resolveProfileImageUrl(href: string): string {
    const t = href.trim();
    if (t.startsWith("http://") || t.startsWith("https://")) return t;
    if (t.startsWith("//")) return `https:${t}`;
    if (t.startsWith("/")) {
      const base = this.regionGtms?.middleTier;
      if (base) {
        try {
          return new URL(t, base).toString();
        } catch {
          return t;
        }
      }
    }
    return t;
  }

  async fetchAuthenticatedImageSrc(imageUrl: string): Promise<string | null> {
    await this.refreshIfNeeded();
    const resolved = this.resolveProfileImageUrl(imageUrl);
    const cached = this.getCachedImagePath
      ? await this.getCachedImagePath(resolved)
      : null;
    if (cached) return filePathToAssetUrl(cached);
    const accept =
      "image/avif,image/webp,image/apng,image/png,image/*,*/*;q=0.8";
    const companion = this.companionHeadersForTeamsMicrosoftCom(resolved);
    const bearer = this.authToken?.token;
    const skype = this.skypeToken;
    const asmSet =
      skype && this.usesAsmImageAuth(resolved)
        ? {
            Authorization: `skype_token ${skype}`,
            Accept: accept,
          }
        : null;
    const bearerSet = bearer
      ? {
          Authorization: `Bearer ${bearer}`,
          Accept: accept,
          ...companion,
        }
      : null;
    const skypeSet = skype
      ? {
          Authentication: `skypetoken=${skype}`,
          Accept: accept,
          ...companion,
        }
      : null;
    const combinedSet =
      bearer && skype
        ? {
            Authorization: `Bearer ${bearer}`,
            Authentication: `skypetoken=${skype}`,
            Accept: accept,
            ...companion,
          }
        : null;
    const headerSets: Record<string, string>[] = [];
    const cdnFirst = this.preferSkypeTokenFirstForImageFetch(resolved);
    if (cdnFirst) {
      if (asmSet) headerSets.push(asmSet);
      if (skypeSet) headerSets.push(skypeSet);
      if (bearerSet) headerSets.push(bearerSet);
      if (combinedSet) headerSets.push(combinedSet);
    } else {
      if (bearerSet) headerSets.push(bearerSet);
      if (asmSet) headerSets.push(asmSet);
      if (skypeSet) headerSets.push(skypeSet);
      if (combinedSet) headerSets.push(combinedSet);
    }
    if (headerSets.length === 0) {
      return null;
    }

    for (const headers of headerSets) {
      const res = await this.httpFetch(resolved, {
        headers,
        redirect: "follow",
      });
      if (!res.ok) continue;
      const buf = await res.arrayBuffer();
      if (buf.byteLength === 0 || buf.byteLength > 20_000_000) return null;
      let ct = res.headers.get("content-type")?.split(";")[0]?.trim() || "";
      if (!ct.startsWith("image/")) {
        const sniffed = sniffImageMimeFromBuffer(buf);
        if (!sniffed) return null;
        ct = sniffed;
      }
      const bytes = new Uint8Array(buf);
      const filePath = await cacheImageFile(
        resolved,
        bytes,
        imageFileExtensionFromMime(ct) ?? undefined,
      );
      if (this.setCachedImagePath) {
        await this.setCachedImagePath(resolved, filePath);
      }
      return filePathToAssetUrl(filePath);
    }
    return null;
  }

  private async fetchWithSkypeToken<T>(url: string): Promise<T> {
    const res = await this.sendWithSkypeToken(url);
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(
        `API request failed (${res.status}): ${body.slice(0, 200)}`,
      );
    }

    return res.json() as Promise<T>;
  }

  private async sendWithSkypeToken(
    url: string,
    init?: RequestInit,
  ): Promise<Response> {
    return this.httpFetch(url, {
      ...init,
      headers: {
        Authentication: `skypetoken=${this.skypeToken}`,
        ...(init?.headers ?? {}),
      },
    });
  }

  private async fetchWithBearer<T>(url: string): Promise<T> {
    const res = await this.httpFetch(url, {
      headers: {
        Authorization: `Bearer ${this.authToken?.token}`,
        ...this.companionHeadersForTeamsMicrosoftCom(url),
      },
    });

    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(
        `API request failed (${res.status}): ${body.slice(0, 200)}`,
      );
    }

    return res.json() as Promise<T>;
  }

  private async resolveAttachmentParticipantIds(
    conversationId: string,
    participantIds?: string[],
  ): Promise<string[]> {
    const directParticipants = normalizeParticipantIds(participantIds ?? []);
    if (directParticipants.length > 0) {
      return directParticipants;
    }

    const conversation = await this.getConversation(conversationId);
    const conversationParticipants = normalizeParticipantIds(
      conversation.members?.map((member) => member.id ?? "") ?? [],
    );
    if (conversationParticipants.length > 0) {
      return conversationParticipants;
    }

    return normalizeParticipantIds([conversationId]);
  }
}
