import { preloadConversationThread } from "@better-teams/app/features/chat/thread/preload";
import {
  useActiveTeamsAccount,
  useMaintainTeamsAvailability,
  useTeamsConversations,
  useTeamsPresence,
  useTeamsProfilePresentation,
  useTeamsSession,
} from "@better-teams/app/features/chat/workspace/teams-hooks";
import {
  beginPerfMeasure,
  countDomNodes,
  isPerfEnabled,
  recordPerfMetric,
  updatePerfSnapshot,
} from "@better-teams/app/platform/perf";
import { teamsKeys } from "@better-teams/app/services/teams/query-keys";
import { canonAvatarMri } from "@better-teams/core/teams/profile/avatars";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  useTransition,
} from "react";
import type { ComposerMentionCandidate } from "../composer/Composer";
import type { ProfileData } from "../profile/ProfileCard";
import type { ThreadViewHandle } from "../thread/ThreadView";
import {
  useConversationHoverPrefetch,
  useFavoriteConversationMutation,
  useSharedConversationLookup,
  useSidebarConversationViewModel,
} from "./workspace-hooks";

type SelectionFocusTarget = "sidebar" | "thread" | "composer";

export function useProductivityWorkspaceController() {
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
  const [selectionFocusTarget, setSelectionFocusTarget] =
    useState<SelectionFocusTarget>("thread");
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
    (id: string, focus: SelectionFocusTarget) => {
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

  const handleSubmitSearch = useCallback((query: string) => {
    const trimmedQuery = query.trim();
    setThreadSearchQuery(trimmedQuery);
    if (!trimmedQuery) return;
    threadViewRef.current?.submitSearch(trimmedQuery);
  }, []);
  const handleCloseSearch = useCallback(() => {
    setThreadSearchQuery("");
    setThreadSearchResultCount(0);
  }, []);
  const handleOpenSelectedProfile = selectedProfileData
    ? () => setProfileSidebarProfile(selectedProfileData)
    : undefined;
  const selectedProfileButtonLabel = selectedProfileData
    ? `Open profile for ${selectedProfileData.displayName}`
    : undefined;

  const accountLoading = !session && sessionQuery.isPending;
  const conversationsLoading =
    allSidebarItems.length === 0 &&
    (!session || conversationsQuery.isPending || conversationsQuery.isFetching);
  const errorMessage =
    !session && sessionQuery.isError
      ? sessionQuery.error instanceof Error
        ? sessionQuery.error.message
        : "Could not connect to Teams"
      : null;

  return {
    activeTenantId,
    accounts,
    accountAvatarByTenant,
    accountLoading,
    activeConversationId,
    allSidebarItems,
    announcement,
    avatarFullByMri,
    avatarThumbByMri,
    companyNameByMri,
    composerMentionCandidates,
    composerRef,
    conversationsLoading,
    departmentByMri,
    displayNameByMri,
    emailByMri,
    errorMessage,
    handleCloseProfileSidebar: () => setProfileSidebarProfile(null),
    handleCloseSearch,
    handleHoverConversation,
    handleHoverConversationEnd,
    handleOpenThreadProfile: setProfileSidebarProfile,
    handleOpenSelectedProfile,
    handleSelectConversation,
    handleSubmitSearch,
    handleToggleFavorite,
    isSwitchingAccount,
    jobTitleByMri,
    liveSessionReady,
    locationByMri,
    pendingSelectedId,
    presenceByMri,
    profileSidebarData,
    searchInputRef,
    selectedItem,
    selectedProfileButtonLabel,
    selectionFocusTarget,
    selfAvatarSrc,
    session,
    sessionQuery,
    setThreadSearchQuery,
    setThreadSearchResultCount,
    sharedConversationsByMri,
    switchAccount,
    tenantNameByMri,
    threadSearchQuery,
    threadSearchResultCount,
    threadViewRef,
    workspaceRef,
  };
}
