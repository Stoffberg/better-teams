import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  useTransition,
} from "react";
import {
  conversationChatKind,
  conversationKindShortLabel,
  conversationPreview,
  conversationTitle,
  formatSidebarTime,
  messageTimestamp,
} from "@/lib/chat-format";
import {
  beginPerfMeasure,
  countDomNodes,
  isPerfEnabled,
  PerfProfiler,
  recordPerfMetric,
  updatePerfSnapshot,
} from "@/lib/perf";
import { buildSharedConversationsByMri } from "@/lib/shared-conversation-index";
import { getOrCreateClient } from "@/lib/teams-client-factory";
import {
  useActiveTeamsAccount,
  useMaintainTeamsAvailability,
  useTeamsConversations,
  useTeamsPresence,
  useTeamsProfilePresentation,
  useTeamsSession,
} from "@/lib/teams-hooks";
import {
  canonAvatarMri,
  dmConversationAvatarMri,
} from "@/lib/teams-profile-avatars";
import { teamsKeys } from "@/lib/teams-query-keys";
import { preloadConversationThread } from "@/lib/teams-thread-preload";
import { useTeamsAccountContext } from "@/providers/TeamsAccountProvider";
import type { Conversation, ConversationMember } from "@/services/teams/types";
import { Skeleton } from "../ui/skeleton";
import type { ComposerMentionCandidate } from "./Composer";
import { Composer } from "./Composer";
import type { ProfileData } from "./ProfileCard";
import { ProfileSidebar } from "./ProfileCard";
import { Sidebar } from "./Sidebar";
import { ThreadHeader } from "./ThreadHeader";
import { ThreadView, type ThreadViewHandle } from "./ThreadView";
import type { SidebarConversationItem } from "./types";

function peerProfileDisplayName(
  conversation: Conversation,
  selfSkypeId: string | undefined,
  byMri: Record<string, string>,
): string | undefined {
  const mri = dmConversationAvatarMri(conversation, selfSkypeId);
  if (!mri) return undefined;
  const displayName = byMri[canonAvatarMri(mri)];
  return typeof displayName === "string" && displayName.trim()
    ? displayName.trim()
    : undefined;
}

function isConversationFavorite(conversation: Conversation): boolean {
  const favorite = conversation.properties?.favorite;
  if (typeof favorite === "boolean") return favorite;
  if (typeof favorite !== "string") return false;
  return favorite.trim().toLowerCase() === "true";
}

function updateConversationFavoriteState(
  conversations: Conversation[],
  conversationId: string,
  favorite: boolean,
): Conversation[] {
  return conversations.map((conversation) =>
    conversation.id === conversationId
      ? {
          ...conversation,
          properties: {
            ...(conversation.properties ?? {}),
            favorite,
          },
        }
      : conversation,
  );
}

function MainLoadingSkeleton() {
  return (
    <div className="flex flex-1 flex-col px-8 py-8">
      <div className="mb-8 flex items-center gap-3">
        <Skeleton className="size-11 rounded-xl" />
        <div className="space-y-2">
          <Skeleton className="h-4 w-44" />
          <Skeleton className="h-3 w-28" />
        </div>
      </div>
      <div className="space-y-6">
        {[0.74, 0.48, 0.62].map((width) => (
          <div key={width} className="flex gap-3">
            <Skeleton className="size-9 shrink-0 rounded-xl" />
            <div className="flex-1 space-y-2">
              <Skeleton className="h-3.5 w-24" />
              <Skeleton
                className="h-10 rounded-xl"
                style={{ width: `${width * 100}%` }}
              />
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function TenantScopedWorkspace() {
  const queryClient = useQueryClient();
  const { activeTenantId, accounts, isSwitchingAccount, switchAccount } =
    useActiveTeamsAccount();
  const sessionQuery = useTeamsSession();
  const session = sessionQuery.session;
  const liveSessionReady = Boolean(session);
  useMaintainTeamsAvailability(liveSessionReady);
  const conversationsQuery = useTeamsConversations(liveSessionReady);
  const conversations = conversationsQuery.conversations;
  const profilePresentation = useTeamsProfilePresentation({
    conversations,
    selfSkypeId: session?.skypeId,
  });
  const presenceByMri = useTeamsPresence({
    conversations,
    selfSkypeId: session?.skypeId,
  });
  const avatarThumbByMri = profilePresentation.avatarThumbs;
  const avatarFullByMri = profilePresentation.avatarFull;
  const displayNameByMri = profilePresentation.displayNames;
  const emailByMri = profilePresentation.emails;
  const jobTitleByMri = profilePresentation.jobTitles;
  const departmentByMri = profilePresentation.departments;
  const companyNameByMri = profilePresentation.companyNames;
  const tenantNameByMri = profilePresentation.tenantNames;
  const locationByMri = profilePresentation.locations;

  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [pendingSelectedId, setPendingSelectedId] = useState<string | null>(
    null,
  );
  const [announcement, setAnnouncement] = useState("");
  const [isProfileSidebarOpen, setIsProfileSidebarOpen] = useState(false);
  const [threadSearchQuery, setThreadSearchQuery] = useState("");
  const [threadSearchResultCount, setThreadSearchResultCount] = useState(0);
  const [selectionFocusTarget, setSelectionFocusTarget] = useState<
    "sidebar" | "thread" | "composer"
  >("thread");
  const perfEnabled = isPerfEnabled();

  const [, startTransition] = useTransition();

  const searchInputRef = useRef<HTMLInputElement>(null);
  const composerRef = useRef<HTMLDivElement>(null);
  const threadViewRef = useRef<ThreadViewHandle>(null);
  const hoverPrefetchTimeoutsRef = useRef<Record<string, number>>({});
  const workspaceRef = useRef<HTMLDivElement>(null);
  const pendingSelectionMeasureRef = useRef<ReturnType<
    typeof beginPerfMeasure
  > | null>(null);
  const activeConversationIdRef = useRef<string | null>(null);

  const activitySortedConversations = useMemo(
    () =>
      [...conversations].sort((left, right) => {
        const rightTimestamp = Date.parse(
          right.lastMessage ? messageTimestamp(right.lastMessage) : "",
        );
        const leftTimestamp = Date.parse(
          left.lastMessage ? messageTimestamp(left.lastMessage) : "",
        );
        const safeRight = Number.isNaN(rightTimestamp) ? 0 : rightTimestamp;
        const safeLeft = Number.isNaN(leftTimestamp) ? 0 : leftTimestamp;
        if (safeRight !== safeLeft) return safeRight - safeLeft;
        return conversationTitle(left, session?.skypeId).localeCompare(
          conversationTitle(right, session?.skypeId),
        );
      }),
    [conversations, session?.skypeId],
  );

  const favoriteMutation = useMutation({
    mutationFn: async ({
      conversationId,
      favorite,
    }: {
      conversationId: string;
      favorite: boolean;
    }) => {
      const client = await getOrCreateClient(activeTenantId);
      await client.setConversationFavorite(conversationId, favorite);
      return { conversationId, favorite };
    },
    onMutate: async ({ conversationId, favorite }) => {
      await queryClient.cancelQueries({
        queryKey: teamsKeys.conversations(activeTenantId),
      });
      const previousConversations =
        queryClient.getQueryData<Conversation[]>(
          teamsKeys.conversations(activeTenantId),
        ) ?? [];
      queryClient.setQueryData<Conversation[]>(
        teamsKeys.conversations(activeTenantId),
        updateConversationFavoriteState(
          previousConversations,
          conversationId,
          favorite,
        ),
      );
      return { previousConversations };
    },
    onError: (_error, _variables, context) => {
      if (!context?.previousConversations) return;
      queryClient.setQueryData(
        teamsKeys.conversations(activeTenantId),
        context.previousConversations,
      );
    },
    onSettled: async () => {
      await queryClient.invalidateQueries({
        queryKey: teamsKeys.conversations(activeTenantId),
      });
    },
  });

  const allSidebarItems = useMemo<SidebarConversationItem[]>(() => {
    const items = activitySortedConversations.flatMap((conversation) => {
      const title = conversationTitle(
        conversation,
        session?.skypeId,
        peerProfileDisplayName(
          conversation,
          session?.skypeId,
          displayNameByMri,
        ),
      );
      const kind = conversationChatKind(conversation);
      if (kind === "dm" && title === "Direct message") return [];
      const preview = conversationPreview(conversation);
      const dmMri = dmConversationAvatarMri(conversation, session?.skypeId);
      const avatarThumbSrc = dmMri
        ? avatarThumbByMri[canonAvatarMri(dmMri)]
        : undefined;
      const avatarFullSrc = dmMri
        ? avatarFullByMri[canonAvatarMri(dmMri)]
        : undefined;
      const lastMessage = conversation.lastMessage;
      const ts = lastMessage ? messageTimestamp(lastMessage) : "";
      return [
        {
          id: conversation.id,
          conversation,
          title,
          preview,
          kind,
          isFavorite: isConversationFavorite(conversation),
          avatarMri: dmMri ? canonAvatarMri(dmMri) : undefined,
          avatarThumbSrc,
          avatarFullSrc,
          sideTime: ts ? formatSidebarTime(ts) : "",
          searchText:
            `${title} ${preview} ${conversationKindShortLabel(kind)}`.toLowerCase(),
        },
      ];
    });

    const favoriteItems = items
      .filter((item) => item.isFavorite)
      .sort((left, right) => left.title.localeCompare(right.title));
    const otherItems = items.filter((item) => !item.isFavorite);

    return [...favoriteItems, ...otherItems];
  }, [
    activitySortedConversations,
    avatarFullByMri,
    avatarThumbByMri,
    displayNameByMri,
    session?.skypeId,
  ]);

  const sidebarItemById = useMemo(
    () => Object.fromEntries(allSidebarItems.map((item) => [item.id, item])),
    [allSidebarItems],
  );
  const sidebarDisplayNameByMri = useMemo(
    () =>
      Object.fromEntries(
        allSidebarItems.flatMap((item) =>
          item.avatarMri ? [[item.avatarMri, item.title] as const] : [],
        ),
      ),
    [allSidebarItems],
  );
  const selectedSidebarItem = selectedId
    ? (sidebarItemById[selectedId] ?? null)
    : null;
  const selectedProfileMri =
    selectedSidebarItem?.kind === "dm" ? selectedSidebarItem.avatarMri : null;
  const sharedConversationCandidateIds = useMemo(
    () =>
      allSidebarItems
        .filter((item) => item.kind !== "dm")
        .map((item) => item.id)
        .sort(),
    [allSidebarItems],
  );
  const sharedConversationDetailsQuery = useQuery({
    queryKey: [
      "teams",
      "shared-thread-members",
      activeTenantId ?? "__default__",
      selectedProfileMri ?? "__none__",
      sharedConversationCandidateIds.join("\x1f"),
    ],
    queryFn: async () => {
      const byConversationId: Record<string, ConversationMember[]> = {};
      const missingConversationIds: string[] = [];

      for (const conversationId of sharedConversationCandidateIds) {
        const cachedMembers = queryClient.getQueryData<ConversationMember[]>(
          teamsKeys.threadMembers(activeTenantId, conversationId),
        );
        if (cachedMembers && cachedMembers.length > 0) {
          byConversationId[conversationId] = cachedMembers;
          continue;
        }

        const sidebarMembers =
          sidebarItemById[conversationId]?.conversation.members;
        if (sidebarMembers && sidebarMembers.length > 0) {
          byConversationId[conversationId] = sidebarMembers;
          queryClient.setQueryData(
            teamsKeys.threadMembers(activeTenantId, conversationId),
            sidebarMembers,
          );
          continue;
        }

        missingConversationIds.push(conversationId);
      }

      if (missingConversationIds.length === 0) {
        return byConversationId;
      }

      const client = await getOrCreateClient(activeTenantId ?? undefined);
      const concurrency = 6;
      let index = 0;

      await Promise.all(
        Array.from(
          { length: Math.min(concurrency, missingConversationIds.length) },
          async () => {
            while (index < missingConversationIds.length) {
              const conversationId = missingConversationIds[index];
              index += 1;
              if (!conversationId) continue;
              try {
                const members = await client.getThreadMembers(conversationId);
                byConversationId[conversationId] = members;
                queryClient.setQueryData(
                  teamsKeys.threadMembers(activeTenantId, conversationId),
                  members,
                );
              } catch {}
            }
          },
        ),
      );

      return byConversationId;
    },
    enabled:
      isProfileSidebarOpen &&
      Boolean(selectedProfileMri) &&
      sharedConversationCandidateIds.length > 0,
    staleTime: 5 * 60_000,
  });
  const detailedSharedConversationById =
    sharedConversationDetailsQuery.data ?? {};
  const sharedConversationsByMri = useMemo(
    () =>
      buildSharedConversationsByMri(
        allSidebarItems,
        detailedSharedConversationById,
        { ...sidebarDisplayNameByMri, ...displayNameByMri },
        emailByMri,
      ),
    [
      allSidebarItems,
      detailedSharedConversationById,
      sidebarDisplayNameByMri,
      displayNameByMri,
      emailByMri,
    ],
  );

  const activeConversationId = useMemo(() => {
    if (selectedId == null) return null;
    return conversations.some((conversation) => conversation.id === selectedId)
      ? selectedId
      : null;
  }, [conversations, selectedId]);
  activeConversationIdRef.current = activeConversationId;
  const openConversationRequest = useQuery({
    queryKey: ["open-conversation-request"],
    queryFn: async () => null as string | null,
    enabled: false,
    initialData: null,
    staleTime: Number.POSITIVE_INFINITY,
    gcTime: Number.POSITIVE_INFINITY,
  });

  const selectedItem = activeConversationId
    ? (sidebarItemById[activeConversationId] ?? null)
    : null;
  const selectedProfileData = useMemo<ProfileData | null>(() => {
    if (
      !selectedItem ||
      selectedItem.kind !== "dm" ||
      !selectedItem.avatarMri
    ) {
      return null;
    }
    const mri = selectedItem.avatarMri;
    return {
      mri,
      displayName: displayNameByMri[mri] || selectedItem.title,
      avatarThumbSrc: avatarThumbByMri[mri],
      avatarFullSrc: avatarFullByMri[mri] ?? avatarThumbByMri[mri],
      email: emailByMri[mri],
      jobTitle: jobTitleByMri[mri],
      department: departmentByMri[mri],
      companyName: companyNameByMri[mri],
      tenantName: tenantNameByMri[mri],
      location: locationByMri[mri],
      presence: presenceByMri[mri],
      onMessage:
        selectedItem.kind === "dm"
          ? undefined
          : () => {
              setSelectedId(selectedItem.id);
              setSelectionFocusTarget("composer");
              setAnnouncement(
                `Ready to message ${displayNameByMri[mri] || selectedItem.title}`,
              );
              setIsProfileSidebarOpen(false);
            },
      onOpenConversation: (conversationId: string) => {
        const item = sidebarItemById[conversationId];
        setSelectedId(conversationId);
        setSelectionFocusTarget("thread");
        setAnnouncement(item ? `Opened ${item.title}` : "Opened conversation");
        setIsProfileSidebarOpen(false);
      },
      currentConversationId: selectedItem.id,
      sharedConversationHeading: `Other chats with ${displayNameByMri[mri] || selectedItem.title}`,
      sharedConversations: sharedConversationsByMri[mri] ?? [],
    };
  }, [
    avatarFullByMri,
    avatarThumbByMri,
    companyNameByMri,
    departmentByMri,
    displayNameByMri,
    emailByMri,
    jobTitleByMri,
    locationByMri,
    presenceByMri,
    selectedItem,
    sidebarItemById,
    sharedConversationsByMri,
    tenantNameByMri,
  ]);
  const composerMentionCandidates = useMemo<ComposerMentionCandidate[]>(() => {
    const seen = new Set<string>();
    const pushCandidate = (
      mriValue: string | undefined,
      displayName: string | undefined,
      email?: string,
    ) => {
      if (!mriValue || !displayName) return [] as ComposerMentionCandidate[];
      const mri = canonAvatarMri(mriValue);
      if (seen.has(mri)) return [] as ComposerMentionCandidate[];
      seen.add(mri);
      return [
        {
          mri,
          displayName,
          email: emailByMri[mri] || email,
          avatarSrc: avatarThumbByMri[mri],
        },
      ];
    };

    const selectedMembers =
      selectedItem?.conversation.members?.flatMap((member) =>
        pushCandidate(
          member.id,
          displayNameByMri[canonAvatarMri(member.id)] ||
            member.displayName ||
            member.friendlyName ||
            member.userPrincipalName,
          member.userPrincipalName,
        ),
      ) ?? [];

    const selectedDmPeer =
      selectedItem?.avatarMri && selectedItem.kind === "dm"
        ? pushCandidate(
            selectedItem.avatarMri,
            displayNameByMri[selectedItem.avatarMri] || selectedItem.title,
            emailByMri[selectedItem.avatarMri],
          )
        : [];

    const sidebarPeers = allSidebarItems.flatMap((item) =>
      item.avatarMri
        ? pushCandidate(
            item.avatarMri,
            displayNameByMri[item.avatarMri] || item.title,
            emailByMri[item.avatarMri],
          )
        : [],
    );

    const profileDirectory = Object.entries(displayNameByMri).flatMap(
      ([mri, displayName]) => pushCandidate(mri, displayName, emailByMri[mri]),
    );

    return [
      ...selectedMembers,
      ...selectedDmPeer,
      ...sidebarPeers,
      ...profileDirectory,
    ];
  }, [
    allSidebarItems,
    avatarThumbByMri,
    displayNameByMri,
    emailByMri,
    selectedItem,
  ]);

  const selfAvatarSrc = session?.skypeId
    ? avatarThumbByMri[
        canonAvatarMri(
          session.skypeId.startsWith("8:")
            ? session.skypeId
            : `8:${session.skypeId}`,
        )
      ]
    : undefined;

  const accountAvatarByTenant = useMemo(
    () =>
      activeTenantId && selfAvatarSrc
        ? { [activeTenantId]: selfAvatarSrc }
        : {},
    [activeTenantId, selfAvatarSrc],
  );
  useEffect(() => {
    const requestedConversationId = openConversationRequest.data;
    if (!requestedConversationId) return;
    const requestedItem = sidebarItemById[requestedConversationId];
    if (!requestedItem) return;
    setSelectedId(requestedConversationId);
    setSelectionFocusTarget("thread");
    queryClient.setQueryData(["open-conversation-request"], null);
  }, [openConversationRequest.data, queryClient, sidebarItemById]);

  useEffect(() => {
    if (selectionFocusTarget !== "composer") return;
    composerRef.current?.focus();
  }, [selectionFocusTarget]);

  useEffect(() => {
    if (!perfEnabled) return;
    updatePerfSnapshot("workspace.sidebar", {
      conversationCount: allSidebarItems.length,
      favoriteCount: allSidebarItems.filter((item) => item.isFavorite).length,
      dmCount: allSidebarItems.filter((item) => item.kind === "dm").length,
      groupCount: allSidebarItems.filter((item) => item.kind === "group")
        .length,
      meetingCount: allSidebarItems.filter((item) => item.kind === "meeting")
        .length,
      domNodeCount: countDomNodes(workspaceRef.current),
      selectedConversation:
        activeConversationId ?? pendingSelectedId ?? "__none__",
    });
  }, [activeConversationId, allSidebarItems, pendingSelectedId, perfEnabled]);

  useEffect(() => {
    if (!activeConversationId || !pendingSelectionMeasureRef.current) return;
    pendingSelectionMeasureRef.current({
      conversationId: activeConversationId,
      focusTarget: selectionFocusTarget,
    });
    pendingSelectionMeasureRef.current = null;
  }, [activeConversationId, selectionFocusTarget]);

  const handleSelectConversation = useCallback(
    (id: string, focus: "sidebar" | "thread" | "composer") => {
      const item = sidebarItemById[id];
      if (liveSessionReady) {
        void queryClient.prefetchQuery({
          queryKey: teamsKeys.thread(activeTenantId, id),
          queryFn: () =>
            preloadConversationThread(activeTenantId ?? undefined, id, 60_000),
          staleTime: 25_000,
        });
      }
      recordPerfMetric("workspace.selectConversation.requested", {
        conversationId: id,
        focusTarget: focus,
      });
      pendingSelectionMeasureRef.current = beginPerfMeasure(
        "workspace.selectConversation",
        {
          conversationId: id,
          focusTarget: focus,
        },
      );
      setPendingSelectedId(id);
      setSelectionFocusTarget(focus);
      startTransition(() => {
        setSelectedId(id);
        setPendingSelectedId(null);
        setThreadSearchQuery("");
        setThreadSearchResultCount(0);
        setAnnouncement(item ? `Opened ${item.title}` : "Opened conversation");
        setIsProfileSidebarOpen(false);
      });
    },
    [activeTenantId, liveSessionReady, queryClient, sidebarItemById],
  );

  useEffect(
    () => () => {
      for (const timeoutId of Object.values(hoverPrefetchTimeoutsRef.current)) {
        window.clearTimeout(timeoutId);
      }
    },
    [],
  );

  const handleHoverConversation = useCallback(
    (conversationId: string) => {
      if (
        !liveSessionReady ||
        conversationId === activeConversationIdRef.current
      ) {
        return;
      }
      if (hoverPrefetchTimeoutsRef.current[conversationId]) return;
      const cachedThreadState = queryClient.getQueryState(
        teamsKeys.thread(activeTenantId, conversationId),
      );
      if (cachedThreadState?.dataUpdatedAt) {
        const ageMs = Date.now() - cachedThreadState.dataUpdatedAt;
        if (ageMs < 60_000) return;
      }
      hoverPrefetchTimeoutsRef.current[conversationId] = window.setTimeout(
        () => {
          delete hoverPrefetchTimeoutsRef.current[conversationId];
          void queryClient
            .prefetchQuery({
              queryKey: teamsKeys.thread(activeTenantId, conversationId),
              queryFn: () =>
                preloadConversationThread(
                  activeTenantId ?? undefined,
                  conversationId,
                  60_000,
                ),
              staleTime: 25_000,
            })
            .catch(() => undefined);
        },
        120,
      );
    },
    [liveSessionReady, activeTenantId, queryClient],
  );
  const handleHoverConversationEnd = useCallback((conversationId: string) => {
    const timeoutId = hoverPrefetchTimeoutsRef.current[conversationId];
    if (!timeoutId) return;
    window.clearTimeout(timeoutId);
    delete hoverPrefetchTimeoutsRef.current[conversationId];
  }, []);

  const { mutate: mutateFavorite } = favoriteMutation;
  const handleToggleFavorite = useCallback(
    (conversationId: string, favorite: boolean) => {
      mutateFavorite({ conversationId, favorite });
    },
    [mutateFavorite],
  );

  const accountLoading = !session && sessionQuery.isPending;
  const conversationsLoading =
    allSidebarItems.length === 0 &&
    (!session || conversationsQuery.isPending || conversationsQuery.isFetching);

  if (!session && sessionQuery.isError) {
    const message =
      sessionQuery.error instanceof Error
        ? sessionQuery.error.message
        : "Could not connect to Teams";
    return (
      <div className="flex h-full flex-1 flex-col items-center justify-center gap-4 bg-background">
        <div className="flex size-12 items-center justify-center rounded-2xl bg-destructive/10">
          <span className="text-lg text-destructive">!</span>
        </div>
        <p className="max-w-sm text-center text-[13px] text-muted-foreground">
          {message}
        </p>
        <button
          type="button"
          onClick={() => void sessionQuery.refetch()}
          className="rounded-xl bg-primary px-4 py-2 text-[13px] font-medium text-primary-foreground transition-colors hover:bg-primary/90"
        >
          Try again
        </button>
      </div>
    );
  }

  return (
    <>
      <div aria-live="polite" className="sr-only">
        {announcement}
      </div>

      <div ref={workspaceRef} className="flex h-full min-h-0 flex-1">
        <PerfProfiler
          id="Sidebar"
          detail={{
            conversationCount: allSidebarItems.length,
            selectedConversation:
              pendingSelectedId ?? activeConversationId ?? "__none__",
          }}
        >
          <Sidebar
            upn={session?.upn}
            selfAvatarSrc={selfAvatarSrc}
            accountAvatarByTenant={accountAvatarByTenant}
            presenceByMri={presenceByMri}
            accounts={accounts}
            activeTenantId={activeTenantId}
            onSwitchAccount={switchAccount}
            switchPending={isSwitchingAccount || sessionQuery.isFetching}
            allSidebarItems={allSidebarItems}
            activeConversationId={pendingSelectedId ?? activeConversationId}
            onSelectConversation={handleSelectConversation}
            onHoverConversationStart={handleHoverConversation}
            onHoverConversationEnd={handleHoverConversationEnd}
            onToggleFavorite={handleToggleFavorite}
            searchInputRef={searchInputRef}
            accountLoading={accountLoading}
            conversationsLoading={conversationsLoading}
          />
        </PerfProfiler>

        <main className="flex min-h-0 min-w-0 flex-1 flex-col bg-background">
          {conversationsLoading ? (
            <MainLoadingSkeleton />
          ) : !selectedItem && allSidebarItems.length === 0 ? (
            <div className="flex flex-1 flex-col items-center justify-center gap-4">
              <div className="flex size-20 items-center justify-center rounded-3xl bg-accent">
                <span className="text-4xl text-muted-foreground/20">💬</span>
              </div>
              <div className="text-center">
                <p className="text-[15px] font-semibold text-muted-foreground/40">
                  No conversations yet
                </p>
                <p className="mt-1 text-[12px] text-muted-foreground/25">
                  Start or receive a Teams chat to see it here
                </p>
              </div>
            </div>
          ) : !selectedItem ? (
            <div className="flex flex-1 flex-col items-center justify-center gap-4">
              <div className="flex size-20 items-center justify-center rounded-3xl bg-accent">
                <span className="text-4xl text-muted-foreground/20">💬</span>
              </div>
              <div className="text-center">
                <p className="text-[15px] font-semibold text-muted-foreground/40">
                  Select a conversation
                </p>
                <p className="mt-1 text-[12px] text-muted-foreground/25">
                  Choose from your chats on the left to get started
                </p>
              </div>
            </div>
          ) : (
            <>
              <ThreadHeader
                title={selectedItem.title}
                kind={selectedItem.kind}
                memberCount={
                  selectedItem.conversation.memberCount ??
                  selectedItem.conversation.members?.length ??
                  (selectedItem.kind === "dm" ? 2 : null)
                }
                avatarMris={
                  selectedItem.conversation.members
                    ?.map((member) => member.id)
                    .filter(Boolean) ?? []
                }
                avatarByMri={avatarThumbByMri}
                presenceByMri={presenceByMri}
                onOpenProfile={
                  selectedProfileData
                    ? () => setIsProfileSidebarOpen(true)
                    : undefined
                }
                profileButtonLabel={
                  selectedProfileData
                    ? `Open profile for ${selectedProfileData.displayName}`
                    : undefined
                }
                searchQuery={threadSearchQuery}
                searchResultCount={threadSearchResultCount}
                onSearchQueryChange={setThreadSearchQuery}
                onSubmitSearch={(query) => {
                  const trimmedQuery = query.trim();
                  setThreadSearchQuery(trimmedQuery);
                  if (!trimmedQuery) return;
                  threadViewRef.current?.submitSearch(trimmedQuery);
                }}
                onCloseSearch={() => {
                  setThreadSearchQuery("");
                  setThreadSearchResultCount(0);
                }}
              />
              <PerfProfiler
                id="ThreadView"
                detail={{
                  conversationId: activeConversationId as string,
                  kind: selectedItem.kind,
                }}
              >
                <ThreadView
                  ref={threadViewRef}
                  key={`thread-${activeTenantId ?? "__default__"}-${activeConversationId}`}
                  tenantId={activeTenantId}
                  conversationId={activeConversationId as string}
                  conversationKind={selectedItem.kind}
                  liveSessionReady={liveSessionReady}
                  autoFocus={selectionFocusTarget === "thread"}
                  searchQuery={threadSearchQuery}
                  consumptionHorizon={
                    selectedItem.conversation.consumptionHorizon
                  }
                  onSearchResultCountChange={setThreadSearchResultCount}
                  selfSkypeId={session?.skypeId}
                  avatarByMri={avatarThumbByMri}
                  avatarFullByMri={avatarFullByMri}
                  displayNameByMri={displayNameByMri}
                  emailByMri={emailByMri}
                  jobTitleByMri={jobTitleByMri}
                  departmentByMri={departmentByMri}
                  companyNameByMri={companyNameByMri}
                  tenantNameByMri={tenantNameByMri}
                  locationByMri={locationByMri}
                  sharedConversationsByMri={sharedConversationsByMri}
                />
              </PerfProfiler>
              <PerfProfiler
                id="Composer"
                detail={{
                  conversationId: activeConversationId as string,
                  mentionCandidateCount: composerMentionCandidates.length,
                }}
              >
                <Composer
                  key={`composer-${activeTenantId ?? "__default__"}-${activeConversationId}`}
                  tenantId={activeTenantId}
                  conversationId={activeConversationId as string}
                  conversationTitle={selectedItem.title}
                  conversationMembers={
                    selectedItem.conversation.members
                      ?.map((member) => member.id ?? "")
                      .filter(Boolean) ?? []
                  }
                  composerRef={composerRef}
                  liveSessionReady={liveSessionReady}
                  mentionCandidates={composerMentionCandidates}
                />
              </PerfProfiler>
            </>
          )}
        </main>

        {isProfileSidebarOpen && selectedProfileData ? (
          <ProfileSidebar
            profile={selectedProfileData}
            closeLabel="Close profile sidebar"
            onClose={() => setIsProfileSidebarOpen(false)}
          />
        ) : null}
      </div>
    </>
  );
}

export function ProductivityWorkspace() {
  const { activeTenantId } = useTeamsAccountContext();

  return <TenantScopedWorkspace key={activeTenantId ?? "__default__"} />;
}
