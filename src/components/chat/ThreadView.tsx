import { useQuery, useQueryClient } from "@tanstack/react-query";
import {
  forwardRef,
  useCallback,
  useDeferredValue,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import {
  formatDetailedTimestamp,
  formatMessageTime,
  formatThreadDayDividerLabel,
  gapBetweenMessages,
  isEditedMessage,
  isSelfMessage,
  type MessageInlinePart,
  messageReadStatus,
  messageReadTimestamp,
  messageRichPartsForDisplay,
  messageTimestamp,
  parseConsumptionHorizon,
} from "@/lib/chat-format";
import {
  countDomNodes,
  isPerfEnabled,
  measurePerfAsync,
  updatePerfSnapshot,
} from "@/lib/perf";
import { SqliteThreadCache } from "@/lib/sqlite-cache";
import { getOrCreateClient } from "@/lib/teams-client-factory";
import {
  useTeamsPresence,
  useTeamsProfilePresentation,
} from "@/lib/teams-hooks";
import { canonAvatarMri } from "@/lib/teams-profile-avatars";
import { teamsKeys } from "@/lib/teams-query-keys";
import {
  sortMessagesByTimestamp,
  type ThreadQueryData,
  threadQueryDataFromResponse,
} from "@/lib/teams-thread-query";
import type {
  Conversation,
  ConversationMember,
  Message,
} from "@/services/teams/types";
import { MessageRow } from "./MessageRow";
import type { ProfileData } from "./ProfileCard";
import {
  type DisplayMessage,
  type MessageBlock,
  OLDER_LOAD_THROTTLE_MS,
  THREAD_PAGE,
} from "./types";

type MemberReadReceipt = {
  mri: string;
  sequenceId: number;
  timestamp: number;
};

const OLDER_PREFETCH_THRESHOLD_PX = 1800;
const OLDER_PREFETCH_ROOT_MARGIN = "2200px 0px 0px 0px";
const SCROLL_RESTORE_EPSILON_PX = 0.75;
const EXPECTED_OLDER_FETCH_MS = 1500;
const MAX_VELOCITY_PREFETCH_BONUS_PX = 6000;

type ScrollRestoreAnchor = {
  messageId: string | null;
  scrollHeight: number;
  scrollTop: number;
  top: number;
};

function selfMriFromSkypeId(skypeId?: string): string | null {
  if (!skypeId) return null;
  const trimmed = skypeId.trim();
  if (!trimmed) return null;
  return canonAvatarMri(trimmed.startsWith("8:") ? trimmed : `8:${trimmed}`);
}

function memberName(
  member: ConversationMember,
  displayNameByMri: Record<string, string>,
): string {
  const mri = canonAvatarMri(member.id);
  return (
    displayNameByMri[mri] ||
    member.displayName?.trim() ||
    member.friendlyName?.trim() ||
    member.userPrincipalName?.trim() ||
    mri
  );
}

function firstVisibleMessageNode(viewport: HTMLElement): HTMLElement | null {
  const viewportTop = viewport.getBoundingClientRect().top;
  const messageNodes =
    viewport.querySelectorAll<HTMLElement>("[data-message-id]");
  for (const node of messageNodes) {
    if (node.getBoundingClientRect().bottom > viewportTop) {
      return node;
    }
  }
  return null;
}

export function captureScrollRestoreAnchor(
  viewport: HTMLElement,
): ScrollRestoreAnchor {
  const anchorNode = firstVisibleMessageNode(viewport);
  const viewportTop = viewport.getBoundingClientRect().top;
  return {
    messageId: anchorNode?.dataset.messageId ?? null,
    scrollHeight: viewport.scrollHeight,
    scrollTop: viewport.scrollTop,
    top: anchorNode ? anchorNode.getBoundingClientRect().top - viewportTop : 0,
  };
}

export function restoreScrollRestoreAnchor(
  viewport: HTMLElement,
  restore: ScrollRestoreAnchor,
): void {
  if (restore.messageId) {
    const messageNodes =
      viewport.querySelectorAll<HTMLElement>("[data-message-id]");
    const anchorNode =
      [...messageNodes].find(
        (node) => node.dataset.messageId === restore.messageId,
      ) ?? null;
    if (anchorNode) {
      const viewportTop = viewport.getBoundingClientRect().top;
      const nextTop = anchorNode.getBoundingClientRect().top - viewportTop;
      const delta = nextTop - restore.top;
      if (Math.abs(delta) <= SCROLL_RESTORE_EPSILON_PX) return;
      viewport.scrollTop += delta;
      return;
    }
  }
  const targetScrollTop =
    viewport.scrollHeight - restore.scrollHeight + restore.scrollTop;
  if (
    Math.abs(targetScrollTop - viewport.scrollTop) <= SCROLL_RESTORE_EPSILON_PX
  ) {
    return;
  }
  viewport.scrollTop = targetScrollTop;
}

export function olderPrefetchThresholdForVelocity(
  upwardVelocityPxPerMs: number,
): number {
  const velocityBonus = Math.min(
    MAX_VELOCITY_PREFETCH_BONUS_PX,
    Math.max(0, upwardVelocityPxPerMs) * EXPECTED_OLDER_FETCH_MS,
  );
  return OLDER_PREFETCH_THRESHOLD_PX + velocityBonus;
}

export function shouldPrefetchOlderMessages(
  scrollTop: number,
  upwardVelocityPxPerMs = 0,
): boolean {
  return scrollTop <= olderPrefetchThresholdForVelocity(upwardVelocityPxPerMs);
}

export type ThreadViewHandle = {
  submitSearch: (query: string) => void;
};

type ThreadViewProps = {
  tenantId?: string | null;
  conversationId: string;
  conversationKind: "dm" | "group" | "meeting";
  liveSessionReady: boolean;
  autoFocus?: boolean;
  searchQuery: string;
  consumptionHorizon?: string;
  onSearchResultCountChange?: (resultCount: number) => void;
  selfSkypeId?: string;
  avatarByMri: Record<string, string>;
  avatarFullByMri: Record<string, string>;
  displayNameByMri: Record<string, string>;
  emailByMri: Record<string, string>;
  jobTitleByMri: Record<string, string>;
  departmentByMri: Record<string, string>;
  companyNameByMri: Record<string, string>;
  tenantNameByMri: Record<string, string>;
  locationByMri: Record<string, string>;
  sharedConversationsByMri: Record<
    string,
    NonNullable<ProfileData["sharedConversations"]>
  >;
};

export function profileMessageConversationId(
  conversationKind: "dm" | "group" | "meeting",
  _conversationId: string,
  sharedConversations: NonNullable<ProfileData["sharedConversations"]>,
): string | undefined {
  if (conversationKind === "dm") return undefined;
  return sharedConversations.find((conversation) => conversation.kind === "dm")
    ?.id;
}

export function mergeThreadSnapshots(
  primary: ThreadQueryData | null | undefined,
  secondary: ThreadQueryData | null | undefined,
): ThreadQueryData | null {
  const left = primary ?? null;
  const right = secondary ?? null;
  if (!left && !right) return null;
  const mergedMessages = sortMessagesByTimestamp([
    ...new Map(
      [...(left?.messages ?? []), ...(right?.messages ?? [])].map((message) => [
        message.id,
        message,
      ]),
    ).values(),
  ]);
  const olderPageUrl = left?.olderPageUrl ?? right?.olderPageUrl ?? null;
  return {
    messages: mergedMessages,
    olderPageUrl,
    moreOlder:
      (left?.moreOlder ?? false) ||
      (right?.moreOlder ?? false) ||
      (olderPageUrl != null && mergedMessages.length > 0),
  };
}

export const ThreadView = forwardRef<ThreadViewHandle, ThreadViewProps>(
  (
    {
      tenantId,
      conversationId,
      conversationKind,
      liveSessionReady,
      autoFocus,
      searchQuery,
      consumptionHorizon,
      onSearchResultCountChange,
      selfSkypeId,
      avatarByMri,
      avatarFullByMri,
      displayNameByMri,
      emailByMri,
      jobTitleByMri,
      departmentByMri,
      companyNameByMri,
      tenantNameByMri,
      locationByMri,
      sharedConversationsByMri,
    },
    ref,
  ) => {
    const queryClient = useQueryClient();
    const deferredSearchQuery = useDeferredValue(searchQuery);

    const [loadingOlder, setLoadingOlder] = useState(false);
    const [highlightedMessageId, setHighlightedMessageId] = useState<
      string | null
    >(null);
    const perfEnabled = isPerfEnabled();

    const viewportRef = useRef<HTMLDivElement>(null);
    const topSentinelRef = useRef<HTMLLIElement>(null);
    const prevLastMessageIdRef = useRef<string | null>(null);
    const scrollRestoreRef = useRef<ScrollRestoreAnchor | null>(null);
    const lastOlderLoadAtRef = useRef(0);
    const lastScrollSampleRef = useRef<{ top: number; time: number } | null>(
      null,
    );
    const loadingOlderRef = useRef(false);
    const scrollFrameRef = useRef<number | null>(null);
    const pendingScrollMessageIdRef = useRef<string | null>(null);
    const highlightTimerRef = useRef<number | null>(null);
    const searchMatchStateRef = useRef<{ query: string; index: number }>({
      query: "",
      index: -1,
    });

    const threadQuery = useQuery({
      queryKey: teamsKeys.thread(tenantId, conversationId),
      queryFn: async () => {
        return measurePerfAsync(
          "thread.fetchMessages",
          { conversationId, pageSize: THREAD_PAGE, tenantId: tenantId ?? null },
          async () => {
            const client = await getOrCreateClient(tenantId ?? undefined);
            const res = await client.getMessages(
              conversationId,
              THREAD_PAGE,
              1,
            );
            return threadQueryDataFromResponse(res);
          },
        );
      },
      enabled: liveSessionReady,
      staleTime: 25_000,
    });
    const cachedThreadQuery = useQuery({
      queryKey: teamsKeys.threadCache(tenantId, conversationId),
      queryFn: async () =>
        SqliteThreadCache.getSnapshot(tenantId ?? undefined, conversationId),
      staleTime: Number.POSITIVE_INFINITY,
      gcTime: Number.POSITIVE_INFINITY,
    });
    const threadMembersQuery = useQuery({
      queryKey: teamsKeys.threadMembers(tenantId, conversationId),
      queryFn: async () =>
        measurePerfAsync(
          "thread.fetchMembers",
          { conversationId, tenantId: tenantId ?? null },
          async () => {
            const client = await getOrCreateClient(tenantId ?? undefined);
            return client.getThreadMembers(conversationId);
          },
        ),
      enabled: liveSessionReady,
      staleTime: 60_000,
    });
    const consumptionHorizonsQuery = useQuery({
      queryKey: teamsKeys.threadConsumptionHorizons(tenantId, conversationId),
      queryFn: async () => {
        const client = await getOrCreateClient(tenantId ?? undefined);
        return client.getMembersConsumptionHorizon(conversationId);
      },
      enabled: liveSessionReady && conversationKind !== "meeting",
      staleTime: 20_000,
      retry: 1,
    });

    const threadData = useMemo(
      () =>
        mergeThreadSnapshots(
          threadQuery.data,
          cachedThreadQuery.data?.data ?? null,
        ),
      [cachedThreadQuery.data?.data, threadQuery.data],
    );
    const rawMessages = threadData?.messages ?? [];
    const threadMembers = threadMembersQuery.data ?? [];
    const threadLoading =
      threadQuery.isPending && !cachedThreadQuery.data?.data;
    const threadHasData = Boolean(threadData);
    const profileConversations = useMemo<Conversation[]>(
      () =>
        threadMembers.length > 0
          ? [
              {
                id: conversationId,
                members: threadMembers,
              } satisfies Conversation,
            ]
          : [],
      [conversationId, threadMembers],
    );
    const threadProfilePresentation = useTeamsProfilePresentation({
      conversations: profileConversations,
      messages: rawMessages,
      selfSkypeId,
    });
    const threadPresenceByMri = useTeamsPresence({
      conversations: profileConversations,
      messages: rawMessages,
      selfSkypeId,
    });
    const mergedAvatarByMri = useMemo(
      () => ({ ...avatarByMri, ...threadProfilePresentation.avatarThumbs }),
      [avatarByMri, threadProfilePresentation.avatarThumbs],
    );
    const mergedAvatarFullByMri = useMemo(
      () => ({ ...avatarFullByMri, ...threadProfilePresentation.avatarFull }),
      [avatarFullByMri, threadProfilePresentation.avatarFull],
    );
    const mergedDisplayNameByMri = useMemo(
      () => ({
        ...displayNameByMri,
        ...threadProfilePresentation.displayNames,
      }),
      [displayNameByMri, threadProfilePresentation.displayNames],
    );
    const mergedEmailByMri = useMemo(
      () => ({ ...emailByMri, ...threadProfilePresentation.emails }),
      [emailByMri, threadProfilePresentation.emails],
    );
    const mergedJobTitleByMri = useMemo(
      () => ({
        ...jobTitleByMri,
        ...threadProfilePresentation.jobTitles,
      }),
      [jobTitleByMri, threadProfilePresentation.jobTitles],
    );
    const mergedDepartmentByMri = useMemo(
      () => ({
        ...departmentByMri,
        ...threadProfilePresentation.departments,
      }),
      [departmentByMri, threadProfilePresentation.departments],
    );
    const mergedCompanyNameByMri = useMemo(
      () => ({
        ...companyNameByMri,
        ...threadProfilePresentation.companyNames,
      }),
      [companyNameByMri, threadProfilePresentation.companyNames],
    );
    const mergedTenantNameByMri = useMemo(
      () => ({
        ...tenantNameByMri,
        ...threadProfilePresentation.tenantNames,
      }),
      [tenantNameByMri, threadProfilePresentation.tenantNames],
    );
    const mergedLocationByMri = useMemo(
      () => ({
        ...locationByMri,
        ...threadProfilePresentation.locations,
      }),
      [locationByMri, threadProfilePresentation.locations],
    );
    const selfMri = useMemo(
      () => selfMriFromSkypeId(selfSkypeId),
      [selfSkypeId],
    );
    const fallbackPeerHorizons = useMemo(() => {
      const h = parseConsumptionHorizon(consumptionHorizon);
      return h ? [h] : [];
    }, [consumptionHorizon]);
    const memberHorizons = useMemo<MemberReadReceipt[]>(() => {
      return (consumptionHorizonsQuery.data?.consumptionhorizons ?? [])
        .map((entry) => {
          const parsed = parseConsumptionHorizon(entry.consumptionhorizon);
          if (!parsed) return null;
          return {
            mri: canonAvatarMri(entry.id),
            sequenceId: parsed.sequenceId,
            timestamp: parsed.timestamp,
          } satisfies MemberReadReceipt;
        })
        .filter((entry): entry is MemberReadReceipt => entry != null);
    }, [consumptionHorizonsQuery.data?.consumptionhorizons]);
    const peerHorizons = useMemo(() => {
      if (memberHorizons.length > 0) {
        return memberHorizons
          .filter((entry) => entry.mri !== selfMri)
          .map((entry) => ({
            sequenceId: entry.sequenceId,
            timestamp: entry.timestamp,
            messageId: entry.mri,
          }));
      }
      return fallbackPeerHorizons;
    }, [fallbackPeerHorizons, memberHorizons, selfMri]);
    const receiptParticipants = useMemo(
      () =>
        threadMembers.filter((member) => canonAvatarMri(member.id) !== selfMri),
      [selfMri, threadMembers],
    );
    const receiptParticipantByMri = useMemo(
      () =>
        Object.fromEntries(
          receiptParticipants.map((member) => [
            canonAvatarMri(member.id),
            member,
          ]),
        ),
      [receiptParticipants],
    );

    const threadDisplayState = useMemo(() => {
      const displayMessages = rawMessages.flatMap((message) => {
        const parts = messageRichPartsForDisplay(message);
        const bodyText = parts?.body.map((part) => part.text).join("") ?? "";
        const quoteText = parts?.quote?.map((part) => part.text).join("") ?? "";
        const attachmentTitles =
          parts?.attachments.map((attachment) => attachment.title).join(" ") ??
          "";
        if (
          !parts ||
          (!bodyText.trim() &&
            !quoteText.trim() &&
            parts.attachments.length === 0 &&
            !message.deleted)
        ) {
          return [];
        }
        const self = isSelfMessage(message.from, selfSkypeId);
        const messageSequenceId =
          message.sequenceId ?? (message.id ? Number(message.id) : Number.NaN);
        const seenBy =
          self &&
          conversationKind === "group" &&
          Number.isFinite(messageSequenceId)
            ? memberHorizons
                .filter(
                  (entry) =>
                    entry.mri !== selfMri &&
                    entry.sequenceId >= messageSequenceId,
                )
                .sort((left, right) => right.timestamp - left.timestamp)
                .map((entry) => {
                  const member = receiptParticipantByMri[entry.mri];
                  return {
                    mri: entry.mri,
                    name: member
                      ? memberName(member, mergedDisplayNameByMri)
                      : mergedDisplayNameByMri[entry.mri] || entry.mri,
                    readAt: formatDetailedTimestamp(
                      new Date(entry.timestamp).toISOString(),
                    ),
                  };
                })
            : [];
        const seenMris = new Set(seenBy.map((entry) => entry.mri));
        const unseenBy =
          self && conversationKind === "group"
            ? receiptParticipants
                .filter((member) => !seenMris.has(canonAvatarMri(member.id)))
                .map((member) => ({
                  mri: canonAvatarMri(member.id),
                  name: memberName(member, mergedDisplayNameByMri),
                }))
            : [];
        return [
          {
            message,
            parts,
            displayName: self
              ? "You"
              : message.senderDisplayName?.trim() ||
                message.imdisplayname?.trim() ||
                "Unknown",
            time: formatMessageTime(messageTimestamp(message)),
            self,
            deleted: Boolean(message.deleted),
            edited: isEditedMessage(message),
            readStatus: self
              ? messageReadStatus(message, peerHorizons)
              : undefined,
            sentAt: self
              ? formatDetailedTimestamp(messageTimestamp(message))
              : "",
            readAt: self
              ? formatDetailedTimestamp(
                  messageReadTimestamp(message, peerHorizons),
                )
              : "",
            receiptScope:
              conversationKind === "dm" ? ("dm" as const) : ("group" as const),
            receiptSeenBy: seenBy,
            receiptUnseenBy: unseenBy,
            bodyPreview: [bodyText, quoteText, attachmentTitles]
              .join(" ")
              .replace(/\s+/g, " ")
              .trim(),
            searchText: [
              self
                ? "You"
                : message.senderDisplayName?.trim() ||
                  message.imdisplayname?.trim() ||
                  "Unknown",
              bodyText,
              quoteText,
              attachmentTitles,
            ]
              .join(" ")
              .toLowerCase(),
          },
        ];
      });
      const messageBlocks: MessageBlock[] = [];
      let lastDay = "";
      let previous: DisplayMessage | undefined;
      for (let i = 0; i < displayMessages.length; i++) {
        const entry = displayMessages[i];
        const ts = messageTimestamp(entry.message);
        const day = ts ? formatThreadDayDividerLabel(ts) : "";
        if (day && day !== lastDay) {
          lastDay = day;
          messageBlocks.push({
            kind: "day",
            label: day,
            key: `day-${i}-${day}`,
          });
        }
        const showMeta =
          !previous ||
          previous.message.from !== entry.message.from ||
          gapBetweenMessages(previous.message, entry.message) > 5 * 60 * 1000;
        messageBlocks.push({
          kind: "msg",
          entry,
          messageIndex: i,
          showMeta,
          key: entry.message.id,
        });
        previous = entry;
      }
      return { displayMessages, messageBlocks };
    }, [
      conversationKind,
      memberHorizons,
      mergedDisplayNameByMri,
      rawMessages,
      receiptParticipantByMri,
      receiptParticipants,
      selfSkypeId,
      selfMri,
      peerHorizons,
    ]);
    const displayMessages = threadDisplayState.displayMessages;
    const messageBlocks = threadDisplayState.messageBlocks;
    const loadedMessageCount = rawMessages.length;

    const mentionProfileForPart = useCallback(
      (part: MessageInlinePart): ProfileData | null => {
        if (part.kind !== "mention" || !part.mentionedMri) return null;
        const mri = canonAvatarMri(part.mentionedMri);
        const messageConversationId = profileMessageConversationId(
          conversationKind,
          conversationId,
          sharedConversationsByMri[mri] ?? [],
        );
        return {
          mri,
          displayName:
            mergedDisplayNameByMri[mri] ||
            part.mentionedDisplayName ||
            part.text.replace(/^@/, ""),
          avatarThumbSrc: mergedAvatarByMri[mri],
          avatarFullSrc: mergedAvatarFullByMri[mri] ?? mergedAvatarByMri[mri],
          email: mergedEmailByMri[mri],
          jobTitle: mergedJobTitleByMri[mri],
          department: mergedDepartmentByMri[mri],
          companyName: mergedCompanyNameByMri[mri],
          tenantName: mergedTenantNameByMri[mri],
          location: mergedLocationByMri[mri],
          presence: threadPresenceByMri[mri],
          onOpenConversation: (targetConversationId: string) => {
            queryClient.setQueryData<string | null>(
              ["open-conversation-request"],
              targetConversationId,
            );
          },
          onMessage: messageConversationId
            ? () => {
                queryClient.setQueryData<string | null>(
                  ["open-conversation-request"],
                  messageConversationId,
                );
              }
            : undefined,
          currentConversationId: conversationId,
          sharedConversationHeading: `Other chats with ${
            mergedDisplayNameByMri[mri] ||
            part.mentionedDisplayName ||
            part.text.replace(/^@/, "")
          }`,
          sharedConversations: sharedConversationsByMri[mri] ?? [],
        };
      },
      [
        conversationId,
        conversationKind,
        queryClient,
        sharedConversationsByMri,
        threadPresenceByMri,
        mergedAvatarByMri,
        mergedAvatarFullByMri,
        mergedCompanyNameByMri,
        mergedDepartmentByMri,
        mergedDisplayNameByMri,
        mergedEmailByMri,
        mergedJobTitleByMri,
        mergedLocationByMri,
        mergedTenantNameByMri,
      ],
    );

    const normalizedSearchQuery = deferredSearchQuery.trim().toLowerCase();

    const findMatchingMessageIds = useCallback(
      (normalizedQuery: string) => {
        if (!normalizedQuery) return [];
        return displayMessages
          .filter((entry) => entry.searchText.includes(normalizedQuery))
          .map((entry) => entry.message.id);
      },
      [displayMessages],
    );

    const matchingMessageIds = useMemo(
      () => findMatchingMessageIds(normalizedSearchQuery),
      [findMatchingMessageIds, normalizedSearchQuery],
    );

    useEffect(() => {
      if (!perfEnabled) return;
      updatePerfSnapshot(`thread:${conversationId}`, {
        rawMessageCount: rawMessages.length,
        displayMessageCount: displayMessages.length,
        blockCount: messageBlocks.length,
        memberCount: threadMembers.length,
        searchMatchCount: matchingMessageIds.length,
        domNodeCount: countDomNodes(viewportRef.current),
        hasCachedData: threadData ? 1 : 0,
        loadingOlder: loadingOlder ? 1 : 0,
      });
    }, [
      conversationId,
      displayMessages.length,
      loadingOlder,
      matchingMessageIds.length,
      messageBlocks.length,
      rawMessages.length,
      threadData,
      threadMembers.length,
      perfEnabled,
    ]);

    const tailMessageId =
      displayMessages.length > 0
        ? (displayMessages[displayMessages.length - 1]?.message.id ?? null)
        : null;

    const scrollToMessage = useCallback((messageId: string) => {
      const viewport = viewportRef.current;
      const target = [
        ...(viewport?.querySelectorAll<HTMLElement>("[data-message-id]") ?? []),
      ].find((node) => node.dataset.messageId === messageId);
      if (!target) return;
      pendingScrollMessageIdRef.current = null;
      setHighlightedMessageId(messageId);
      if (highlightTimerRef.current != null) {
        window.clearTimeout(highlightTimerRef.current);
      }
      highlightTimerRef.current = window.setTimeout(() => {
        setHighlightedMessageId((current) =>
          current === messageId ? null : current,
        );
        highlightTimerRef.current = null;
      }, 2200);
      requestAnimationFrame(() => {
        target.scrollIntoView({ block: "center" });
      });
    }, []);

    const runSearch = useCallback(
      (query: string) => {
        const normalizedQuery = query.trim().toLowerCase();
        if (!normalizedQuery) return;
        const immediateMatches = findMatchingMessageIds(normalizedQuery);
        if (immediateMatches.length === 0) return;
        const nextIndex =
          searchMatchStateRef.current.query === normalizedQuery
            ? (searchMatchStateRef.current.index + 1) % immediateMatches.length
            : 0;
        searchMatchStateRef.current = {
          query: normalizedQuery,
          index: nextIndex,
        };
        const messageId = immediateMatches[nextIndex];
        if (messageId) {
          scrollToMessage(messageId);
        }
      },
      [findMatchingMessageIds, scrollToMessage],
    );

    useImperativeHandle(
      ref,
      () => ({
        submitSearch(query: string) {
          runSearch(query);
        },
      }),
      [runSearch],
    );

    const mergeThreadData = useCallback(
      (incoming: ThreadQueryData) => {
        queryClient.setQueryData<ThreadQueryData>(
          teamsKeys.thread(tenantId, conversationId),
          (old) => {
            if (!old) return incoming;
            const merged = new Map(
              old.messages.map((message) => [message.id, message]),
            );
            for (const message of incoming.messages) {
              merged.set(message.id, message);
            }
            const messages = sortMessagesByTimestamp([...merged.values()]);
            return {
              messages,
              olderPageUrl: old.olderPageUrl ?? incoming.olderPageUrl,
              moreOlder: old.moreOlder || incoming.moreOlder,
            };
          },
        );
      },
      [conversationId, queryClient, tenantId],
    );

    const openMessageRef = useCallback(
      async (targetConversationId: string, messageId: string) => {
        if (targetConversationId !== conversationId) {
          queryClient.setQueryData<string | null>(
            ["open-conversation-request"],
            targetConversationId,
          );
          return;
        }
        if (rawMessages.some((message) => message.id === messageId)) {
          scrollToMessage(messageId);
          return;
        }
        const client = await getOrCreateClient(tenantId ?? undefined);
        const res = await client.getAnchoredMessages(
          targetConversationId,
          messageId,
        );
        if (!res) return;
        pendingScrollMessageIdRef.current = messageId;
        mergeThreadData(threadQueryDataFromResponse(res));
      },
      [
        conversationId,
        mergeThreadData,
        queryClient,
        rawMessages,
        scrollToMessage,
        tenantId,
      ],
    );

    const handleDeleteMessage = useCallback(
      async (targetConversationId: string, messageId: string) => {
        try {
          const client = await getOrCreateClient(tenantId ?? undefined);
          await client.deleteMessage(targetConversationId, messageId);
          // Optimistically update local state: mark message as deleted
          queryClient.setQueryData<ThreadQueryData>(
            teamsKeys.thread(tenantId, conversationId),
            (old) => {
              if (!old) return old;
              return {
                ...old,
                messages: old.messages.map((m) =>
                  m.id === messageId
                    ? ({
                        ...m,
                        deleted: true,
                        content: "",
                        properties: {
                          ...m.properties,
                          deletetime: Date.now(),
                        },
                      } as Message)
                    : m,
                ),
              };
            },
          );
        } catch (err) {
          console.error("Failed to delete message:", err);
        }
      },
      [conversationId, queryClient, tenantId],
    );

    const loadOlderMessages = useCallback(async () => {
      if (loadingOlderRef.current) return;
      const snapshot = queryClient.getQueryData<ThreadQueryData>(
        teamsKeys.thread(tenantId, conversationId),
      );
      if (!snapshot?.moreOlder || !snapshot.olderPageUrl) return;
      const now = Date.now();
      if (now - lastOlderLoadAtRef.current < OLDER_LOAD_THROTTLE_MS) return;
      lastOlderLoadAtRef.current = now;
      loadingOlderRef.current = true;
      setLoadingOlder(true);
      try {
        const client = await getOrCreateClient(tenantId ?? undefined);
        const res = await client.getMessagesByUrl(snapshot.olderPageUrl);
        const batch = res.messages ?? [];
        const el = viewportRef.current;
        scrollRestoreRef.current = el ? captureScrollRestoreAnchor(el) : null;
        queryClient.setQueryData<ThreadQueryData>(
          teamsKeys.thread(tenantId, conversationId),
          (old) => {
            if (!old) return old;
            const merged = new Map(old.messages.map((m) => [m.id, m]));
            let added = 0;
            for (const m of batch) {
              if (!merged.has(m.id)) added++;
              merged.set(m.id, m);
            }
            if (added === 0) return { ...old, moreOlder: false };
            const messages = sortMessagesByTimestamp([...merged.values()]);
            const nextUrl = res._metadata?.backwardLink ?? null;
            return {
              messages,
              olderPageUrl: nextUrl,
              moreOlder: nextUrl != null,
            };
          },
        );
      } catch {
        scrollRestoreRef.current = null;
      } finally {
        loadingOlderRef.current = false;
        setLoadingOlder(false);
      }
    }, [conversationId, queryClient, tenantId]);

    useEffect(() => {
      const sentinel = topSentinelRef.current;
      const viewport = viewportRef.current;
      if (!sentinel || !viewport) return;
      const observer = new IntersectionObserver(
        ([entry]) => {
          if (entry?.isIntersecting) void loadOlderMessages();
        },
        {
          root: viewport,
          rootMargin: OLDER_PREFETCH_ROOT_MARGIN,
          threshold: 0,
        },
      );
      observer.observe(sentinel);
      return () => observer.disconnect();
    }, [loadOlderMessages]);

    const onScroll = useCallback(() => {
      const now = performance.now();
      const viewport = viewportRef.current;
      let upwardVelocityPxPerMs = 0;
      if (viewport) {
        const lastSample = lastScrollSampleRef.current;
        if (lastSample) {
          const deltaTime = Math.max(1, now - lastSample.time);
          const deltaTop = lastSample.top - viewport.scrollTop;
          upwardVelocityPxPerMs = deltaTop > 0 ? deltaTop / deltaTime : 0;
        }
        lastScrollSampleRef.current = {
          top: viewport.scrollTop,
          time: now,
        };
      }
      if (scrollFrameRef.current != null) return;
      scrollFrameRef.current = window.requestAnimationFrame(() => {
        scrollFrameRef.current = null;
        const el = viewportRef.current;
        if (!el) return;
        if (!shouldPrefetchOlderMessages(el.scrollTop, upwardVelocityPxPerMs)) {
          return;
        }
        void loadOlderMessages();
      });
    }, [loadOlderMessages]);

    useLayoutEffect(() => {
      if (threadLoading && !threadHasData) return;
      const el = viewportRef.current;
      if (!el) return;
      const last = tailMessageId;
      const prev = prevLastMessageIdRef.current;
      prevLastMessageIdRef.current = last;
      if (last === null) return;
      if (prev === null) {
        el.scrollTop = el.scrollHeight;
        return;
      }
      if (prev !== last) {
        const atBottom = el.scrollHeight - el.scrollTop - el.clientHeight < 150;
        if (atBottom) {
          el.scrollTop = el.scrollHeight;
        }
      }
    }, [threadHasData, threadLoading, tailMessageId]);

    useLayoutEffect(() => {
      const el = viewportRef.current;
      if (!el) return;
      if (loadedMessageCount === 0) return;
      const restore = scrollRestoreRef.current;
      if (!restore) return;
      scrollRestoreRef.current = null;
      restoreScrollRestoreAnchor(el, restore);
    }, [loadedMessageCount]);

    useEffect(() => {
      if (loadingOlderRef.current) return;
      const el = viewportRef.current;
      if (loadedMessageCount === 0) return;
      if (!el || !shouldPrefetchOlderMessages(el.scrollTop)) return;
      const snapshot = queryClient.getQueryData<ThreadQueryData>(
        teamsKeys.thread(tenantId, conversationId),
      );
      if (!snapshot?.moreOlder || !snapshot.olderPageUrl) return;
      const frame = window.requestAnimationFrame(() => {
        void loadOlderMessages();
      });
      return () => window.cancelAnimationFrame(frame);
    }, [
      conversationId,
      loadOlderMessages,
      loadedMessageCount,
      queryClient,
      tenantId,
    ]);

    useLayoutEffect(() => {
      const pendingId = pendingScrollMessageIdRef.current;
      if (!pendingId) return;
      if (!rawMessages.some((message) => message.id === pendingId)) return;
      scrollToMessage(pendingId);
    }, [rawMessages, scrollToMessage]);

    useEffect(() => {
      if (!threadQuery.data || loadingOlder) return;
      let cancelled = false;
      let timer: ReturnType<typeof setTimeout> | null = null;
      let idleHandle: number | null = null;

      const storeThreadSnapshot = () => {
        if (cancelled) return;
        void SqliteThreadCache.storeThread(
          tenantId ?? undefined,
          conversationId,
          threadQuery.data,
        );
      };

      timer = setTimeout(() => {
        if (typeof window !== "undefined" && "requestIdleCallback" in window) {
          idleHandle = window.requestIdleCallback(
            () => {
              idleHandle = null;
              storeThreadSnapshot();
            },
            { timeout: 5_000 },
          );
          return;
        }
        storeThreadSnapshot();
      }, 10_000);

      return () => {
        cancelled = true;
        if (timer != null) {
          clearTimeout(timer);
        }
        if (
          idleHandle != null &&
          typeof window !== "undefined" &&
          "cancelIdleCallback" in window
        ) {
          window.cancelIdleCallback(idleHandle);
        }
      };
    }, [conversationId, loadingOlder, tenantId, threadQuery.data]);

    useEffect(() => {
      onSearchResultCountChange?.(
        normalizedSearchQuery ? matchingMessageIds.length : 0,
      );
    }, [
      matchingMessageIds.length,
      normalizedSearchQuery,
      onSearchResultCountChange,
    ]);

    useEffect(() => {
      return () => {
        if (scrollFrameRef.current != null) {
          window.cancelAnimationFrame(scrollFrameRef.current);
        }
        if (highlightTimerRef.current != null) {
          window.clearTimeout(highlightTimerRef.current);
        }
      };
    }, []);

    useEffect(() => {
      if (normalizedSearchQuery) return;
      searchMatchStateRef.current = { query: "", index: -1 };
      setHighlightedMessageId(null);
    }, [normalizedSearchQuery]);

    useEffect(() => {
      if (!autoFocus) return;
      viewportRef.current?.focus({ preventScroll: true });
    }, [autoFocus]);

    return (
      <section
        ref={viewportRef}
        onScroll={onScroll}
        tabIndex={-1}
        className="relative flex-1 overflow-y-auto overflow-x-hidden bg-background outline-none overscroll-y-contain"
        aria-label="Message thread"
        style={{ overflowAnchor: "none" }}
      >
        {loadingOlder ? (
          <div className="pointer-events-none absolute top-3 left-1/2 z-10 -translate-x-1/2">
            <span className="inline-flex items-center gap-2 rounded-full border border-border bg-background/95 px-4 py-1.5 text-[11px] font-medium text-muted-foreground">
              <span className="size-1.5 animate-pulse rounded-full bg-primary/60" />
              Loading older messages…
            </span>
          </div>
        ) : null}
        <div className="bg-background px-0 pt-4 pb-4">
          {threadQuery.isError ? (
            <div className="flex flex-col items-center gap-4 py-24">
              <div className="flex size-12 items-center justify-center rounded-2xl bg-destructive/10">
                <span className="text-xl">!</span>
              </div>
              <p className="text-[14px] text-muted-foreground">
                Could not load messages
              </p>
              <button
                type="button"
                onClick={() => void threadQuery.refetch()}
                className="rounded-xl bg-primary px-4 py-2 text-[13px] font-medium text-primary-foreground transition-colors hover:bg-primary/90"
              >
                Try again
              </button>
            </div>
          ) : threadLoading && !threadHasData ? (
            <div className="space-y-6 py-10">
              {[0.9, 0.55, 0.75].map((w) => (
                <div key={`skel-${w}`} className="flex gap-3 px-3">
                  <div className="size-9 shrink-0 animate-pulse rounded-xl bg-accent" />
                  <div className="flex-1 space-y-2">
                    <div className="h-3.5 w-24 animate-pulse rounded-lg bg-accent" />
                    <div
                      className="h-10 animate-pulse rounded-xl bg-accent"
                      style={{ width: `${w * 100}%` }}
                    />
                  </div>
                </div>
              ))}
            </div>
          ) : rawMessages.length === 0 ? (
            <div className="flex flex-col items-center gap-3 py-24">
              <div className="flex size-16 items-center justify-center rounded-2xl bg-accent">
                <span className="text-2xl text-muted-foreground/30">💬</span>
              </div>
              <p className="text-[14px] text-muted-foreground/50">
                No messages yet
              </p>
            </div>
          ) : displayMessages.length === 0 ? (
            <p className="py-24 text-center text-[14px] text-muted-foreground/40">
              Only meeting and call activity in this thread.
            </p>
          ) : (
            <ul className="flex flex-col" aria-label="Loaded messages">
              <li
                ref={topSentinelRef}
                className="h-px shrink-0 list-none"
                aria-hidden
              />
              {messageBlocks.map((block) =>
                block.kind === "day" ? (
                  <li key={block.key} className="list-none px-3 py-4">
                    <div className="flex items-center gap-4">
                      <div className="h-px flex-1 bg-border" />
                      <span className="rounded-full bg-accent px-3 py-1 text-[11px] font-semibold tracking-wide text-muted-foreground/60 uppercase">
                        {block.label}
                      </span>
                      <div className="h-px flex-1 bg-border" />
                    </div>
                  </li>
                ) : (
                  (() => {
                    const mri = canonAvatarMri(
                      block.entry.message.fromMri || block.entry.message.from,
                    );
                    const messageConversationId = profileMessageConversationId(
                      conversationKind,
                      conversationId,
                      sharedConversationsByMri[mri] ?? [],
                    );
                    const profile: ProfileData | null = block.entry.self
                      ? null
                      : {
                          mri,
                          displayName:
                            mergedDisplayNameByMri[mri] ||
                            block.entry.displayName,
                          avatarThumbSrc: mergedAvatarByMri[mri],
                          avatarFullSrc:
                            mergedAvatarFullByMri[mri] ??
                            mergedAvatarByMri[mri],
                          email: mergedEmailByMri[mri],
                          jobTitle: mergedJobTitleByMri[mri],
                          department: mergedDepartmentByMri[mri],
                          companyName: mergedCompanyNameByMri[mri],
                          tenantName: mergedTenantNameByMri[mri],
                          location: mergedLocationByMri[mri],
                          presence: threadPresenceByMri[mri],
                          onOpenConversation: (
                            targetConversationId: string,
                          ) => {
                            queryClient.setQueryData<string | null>(
                              ["open-conversation-request"],
                              targetConversationId,
                            );
                          },
                          onMessage: messageConversationId
                            ? () => {
                                queryClient.setQueryData<string | null>(
                                  ["open-conversation-request"],
                                  messageConversationId,
                                );
                              }
                            : undefined,
                          currentConversationId: conversationId,
                          sharedConversationHeading: `Other chats with ${
                            mergedDisplayNameByMri[mri] ||
                            block.entry.displayName
                          }`,
                          sharedConversations:
                            sharedConversationsByMri[mri] ?? [],
                        };
                    return (
                      <MessageRow
                        key={block.key}
                        entry={block.entry}
                        showMeta={block.showMeta}
                        avatarSrc={mergedAvatarByMri[mri]}
                        presence={threadPresenceByMri[mri]}
                        profile={profile}
                        isHighlighted={
                          highlightedMessageId === block.entry.message.id
                        }
                        tenantId={tenantId}
                        onOpenMessageRef={openMessageRef}
                        onDeleteMessage={handleDeleteMessage}
                        getMentionProfile={mentionProfileForPart}
                      />
                    );
                  })()
                ),
              )}
              <li className="h-px shrink-0 list-none" />
            </ul>
          )}
        </div>
      </section>
    );
  },
);
