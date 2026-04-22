import {
  beginPerfMeasure,
  countDomNodes,
  isPerfEnabled,
  PerfProfiler,
  recordPerfMetric,
  updatePerfSnapshot,
} from "@better-teams/app/lib/perf";
import {
  useActiveTeamsAccount,
  useMaintainTeamsAvailability,
  useTeamsConversations,
  useTeamsPresence,
  useTeamsProfilePresentation,
  useTeamsSession,
} from "@better-teams/app/lib/teams-hooks";
import { teamsKeys } from "@better-teams/app/lib/teams-query-keys";
import { preloadConversationThread } from "@better-teams/app/lib/teams-thread-preload";
import { useTeamsAccountContext } from "@better-teams/app/providers/TeamsAccountProvider";
import { canonAvatarMri } from "@better-teams/core/teams-profile-avatars";
import { Skeleton } from "@better-teams/ui/components/skeleton";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  useTransition,
} from "react";
import type { ComposerMentionCandidate } from "./Composer";
import { Composer } from "./Composer";
import type { ProfileData } from "./ProfileCard";
import { ProfileSidebar } from "./ProfileCard";
import { Sidebar } from "./Sidebar";
import { ThreadHeader } from "./ThreadHeader";
import { ThreadView, type ThreadViewHandle } from "./ThreadView";
import {
  useConversationHoverPrefetch,
  useFavoriteConversationMutation,
  useSharedConversationLookup,
  useSidebarConversationViewModel,
} from "./workspace-hooks";

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
  const [profileSidebarProfile, setProfileSidebarProfile] =
    useState<ProfileData | null>(null);
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
  const workspaceRef = useRef<HTMLDivElement>(null);
  const pendingSelectionMeasureRef = useRef<ReturnType<
    typeof beginPerfMeasure
  > | null>(null);

  const { allSidebarItems, sidebarItemById, sidebarDisplayNameByMri } =
    useSidebarConversationViewModel({
      conversations,
      selfSkypeId: session?.skypeId,
      avatarThumbByMri,
      avatarFullByMri,
      displayNameByMri,
    });
  const favoriteMutation = useFavoriteConversationMutation(
    queryClient,
    activeTenantId,
  );
  const profileSidebarMri = profileSidebarProfile?.mri ?? null;
  const sharedConversationsByMri = useSharedConversationLookup({
    activeTenantId,
    profileSidebarMri,
    allSidebarItems,
    sidebarItemById,
    sidebarDisplayNameByMri,
    displayNameByMri,
    emailByMri,
    queryClient,
  });

  const activeConversationId = useMemo(() => {
    if (selectedId == null) return null;
    return conversations.some((conversation) => conversation.id === selectedId)
      ? selectedId
      : null;
  }, [conversations, selectedId]);
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
              setProfileSidebarProfile(null);
            },
      onOpenConversation: (conversationId: string) => {
        const item = sidebarItemById[conversationId];
        setSelectedId(conversationId);
        setSelectionFocusTarget("thread");
        setAnnouncement(item ? `Opened ${item.title}` : "Opened conversation");
        setProfileSidebarProfile(null);
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
  const profileSidebarData = useMemo<ProfileData | null>(() => {
    if (!profileSidebarProfile) return null;
    const mri = profileSidebarProfile.mri;
    const displayName =
      displayNameByMri[mri] || profileSidebarProfile.displayName;
    const onOpenConversation = profileSidebarProfile.onOpenConversation
      ? (conversationId: string) => {
          profileSidebarProfile.onOpenConversation?.(conversationId);
          setProfileSidebarProfile(null);
        }
      : undefined;
    const onMessage = profileSidebarProfile.onMessage
      ? () => {
          profileSidebarProfile.onMessage?.();
          setProfileSidebarProfile(null);
        }
      : undefined;

    return {
      ...profileSidebarProfile,
      displayName,
      avatarThumbSrc:
        avatarThumbByMri[mri] ?? profileSidebarProfile.avatarThumbSrc,
      avatarFullSrc:
        avatarFullByMri[mri] ??
        avatarThumbByMri[mri] ??
        profileSidebarProfile.avatarFullSrc,
      email: emailByMri[mri] ?? profileSidebarProfile.email,
      jobTitle: jobTitleByMri[mri] ?? profileSidebarProfile.jobTitle,
      department: departmentByMri[mri] ?? profileSidebarProfile.department,
      companyName: companyNameByMri[mri] ?? profileSidebarProfile.companyName,
      tenantName: tenantNameByMri[mri] ?? profileSidebarProfile.tenantName,
      location: locationByMri[mri] ?? profileSidebarProfile.location,
      presence: presenceByMri[mri] ?? profileSidebarProfile.presence,
      onOpenConversation,
      onMessage,
      sharedConversationHeading: `Other chats with ${displayName}`,
      sharedConversations:
        sharedConversationsByMri[mri] ??
        profileSidebarProfile.sharedConversations ??
        [],
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
    profileSidebarProfile,
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
        setProfileSidebarProfile(null);
      });
    },
    [activeTenantId, liveSessionReady, queryClient, sidebarItemById],
  );

  const { handleHoverConversation, handleHoverConversationEnd } =
    useConversationHoverPrefetch({
      activeTenantId,
      liveSessionReady,
      activeConversationId,
      queryClient,
    });

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
                    ? () => setProfileSidebarProfile(selectedProfileData)
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
                  onOpenProfile={setProfileSidebarProfile}
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

        {profileSidebarData ? (
          <ProfileSidebar
            profile={profileSidebarData}
            closeLabel="Close profile sidebar"
            onClose={() => setProfileSidebarProfile(null)}
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
