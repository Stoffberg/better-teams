import { getCachedPresence } from "@better-teams/app/lib/electron-bridge";
import { teamsKeys } from "@better-teams/app/lib/teams-query-keys";
import { useTeamsAccountContext } from "@better-teams/app/providers/TeamsAccountProvider";
import { TeamsProfilePresentationSchema } from "@better-teams/core/teams/schemas";
import type {
  Conversation,
  Message,
  PresenceInfo,
  TeamsProfilePresentation,
  TeamsSessionInfo,
} from "@better-teams/core/teams/types";
import { getOrCreateClient } from "@better-teams/core/teams-client-factory";
import {
  canonAvatarMri,
  collectProfileAvatarMris,
} from "@better-teams/core/teams-profile-avatars";
import { useQuery } from "@tanstack/react-query";
import {
  useDeferredValue,
  useEffect,
  useMemo,
  useState,
  useSyncExternalStore,
} from "react";

const PRIORITY_AVATAR_CONVERSATIONS = 12;
const PRESENCE_BATCH_SIZE = 50;
const SELF_AVAILABILITY_HEARTBEAT_MS = 60_000;
const PROFILE_QUERY_GC_MS = 5 * 60_000;
const RESUME_COOLDOWN_MS = 3_000;
const EMPTY_CONVERSATIONS: Conversation[] = [];
const EMPTY_MESSAGES: Message[] = [];
const EMPTY_PRESENCE_BY_MRI: Record<string, PresenceInfo> = {};

function uniqueMris(mris: string[]): string[] {
  const unique = new Map<string, string>();
  for (const mri of mris) {
    const trimmed = mri.trim();
    if (!trimmed) continue;
    unique.set(canonAvatarMri(trimmed), trimmed);
  }
  return [...unique.values()];
}

function isDocumentVisible(): boolean {
  return (
    typeof document === "undefined" || document.visibilityState === "visible"
  );
}

function subscribeToDocumentVisibility(onStoreChange: () => void): () => void {
  if (typeof document === "undefined") return () => undefined;
  document.addEventListener("visibilitychange", onStoreChange);
  return () => document.removeEventListener("visibilitychange", onStoreChange);
}

function useDocumentVisibility(): boolean {
  return useSyncExternalStore(
    subscribeToDocumentVisibility,
    isDocumentVisible,
    () => true,
  );
}

function useResumeCooldown(delayMs = RESUME_COOLDOWN_MS): boolean {
  const documentVisible = useDocumentVisibility();
  const [resumeReady, setResumeReady] = useState(documentVisible);

  useEffect(() => {
    if (!documentVisible) {
      setResumeReady(false);
      return;
    }
    const timer = window.setTimeout(() => {
      setResumeReady(true);
    }, delayMs);
    return () => window.clearTimeout(timer);
  }, [delayMs, documentVisible]);

  return resumeReady;
}

function normalizeProfilePresentation(data: unknown): TeamsProfilePresentation {
  const parsed = TeamsProfilePresentationSchema.safeParse(data);
  if (parsed.success) return parsed.data;
  const raw =
    data && typeof data === "object" ? (data as Record<string, unknown>) : {};
  const record = (value: unknown): Record<string, string> => {
    if (!value || typeof value !== "object" || Array.isArray(value)) return {};
    return Object.fromEntries(
      Object.entries(value).filter(
        (entry): entry is [string, string] => typeof entry[1] === "string",
      ),
    );
  };
  return {
    avatarThumbs: record(raw.avatarThumbs ?? raw.avatars),
    avatarFull: record(raw.avatarFull ?? raw.avatars),
    displayNames: record(raw.displayNames),
    emails: record(raw.emails),
    jobTitles: record(raw.jobTitles),
    departments: record(raw.departments),
    companyNames: record(raw.companyNames),
    tenantNames: record(raw.tenantNames),
    locations: record(raw.locations),
  };
}

export function useActiveTeamsAccount() {
  return useTeamsAccountContext();
}

export function useTeamsSession(): {
  tenantId?: string;
  session?: TeamsSessionInfo;
  isPending: boolean;
  isFetching: boolean;
  isError: boolean;
  error: Error | null;
  refetch: () => Promise<unknown>;
} {
  const { activeTenantId, activeSession } = useTeamsAccountContext();
  const query = useQuery({
    queryKey: teamsKeys.session(activeTenantId),
    queryFn: async () => {
      const client = await getOrCreateClient(activeTenantId);
      const a = client.account;
      return {
        upn: a.upn,
        tenantId: a.tenantId ?? "__default__",
        skypeId: a.skypeId,
        expiresAt: a.expiresAt?.toISOString() ?? null,
        region: a.region,
      } satisfies TeamsSessionInfo;
    },
    initialData: activeSession,
    initialDataUpdatedAt: activeSession ? 0 : undefined,
    enabled: true,
    staleTime: 30_000,
    gcTime: Number.POSITIVE_INFINITY,
  });

  return {
    tenantId: activeTenantId,
    session: query.data ?? activeSession,
    isPending: query.isPending,
    isFetching: query.isFetching,
    isError: query.isError,
    error: query.error instanceof Error ? query.error : null,
    refetch: query.refetch,
  };
}

export function useTeamsConversations(liveSessionReady: boolean): {
  tenantId?: string;
  conversations: Conversation[];
  isPending: boolean;
  isFetching: boolean;
  isError: boolean;
  isSuccess: boolean;
  refetch: () => Promise<unknown>;
} {
  const { activeTenantId } = useTeamsAccountContext();
  const documentVisible = useDocumentVisibility();
  const resumeReady = useResumeCooldown();

  const query = useQuery({
    queryKey: teamsKeys.conversations(activeTenantId),
    queryFn: async () => {
      const client = await getOrCreateClient(activeTenantId);
      const conversationsResponse = await client.getAllConversations(100);
      return conversationsResponse.conversations ?? [];
    },
    enabled: liveSessionReady,
    staleTime: 30_000,
    refetchInterval: () =>
      liveSessionReady && documentVisible && resumeReady ? 30_000 : false,
    refetchIntervalInBackground: false,
  });

  return {
    tenantId: activeTenantId,
    conversations: query.data ?? EMPTY_CONVERSATIONS,
    isPending: query.isPending,
    isFetching: query.isFetching,
    isError: query.isError,
    isSuccess: query.isSuccess,
    refetch: query.refetch,
  };
}

async function fetchProfiles(
  tenantId: string | undefined,
  mris: string[],
): Promise<TeamsProfilePresentation> {
  if (mris.length === 0) return normalizeProfilePresentation(null);
  const client = await getOrCreateClient(tenantId);
  return client.fetchProfileAvatarDataUrls(mris);
}

export function useTeamsProfilePresentation(args: {
  conversations: Conversation[];
  messages?: Message[];
  selfSkypeId?: string;
}) {
  const { activeTenantId } = useTeamsAccountContext();
  const deferredConversations = useDeferredValue(args.conversations);
  const deferredMessages = useDeferredValue(args.messages ?? EMPTY_MESSAGES);
  const profileMris = useMemo(
    () =>
      collectProfileAvatarMris({
        conversations: deferredConversations,
        messages: deferredMessages,
        selfSkypeId: args.selfSkypeId,
      }),
    [deferredConversations, deferredMessages, args.selfSkypeId],
  );
  const priorityProfileMris = useMemo(
    () => profileMris.slice(0, PRIORITY_AVATAR_CONVERSATIONS),
    [profileMris],
  );
  const profileMriSignature = useMemo(
    () => [...profileMris].sort().join("\x1f"),
    [profileMris],
  );
  const prioritySignature = useMemo(
    () => [...priorityProfileMris].sort().join("\x1f"),
    [priorityProfileMris],
  );

  const priorityAvatarQuery = useQuery({
    queryKey: teamsKeys.profileAvatars(activeTenantId, prioritySignature),
    queryFn: () => fetchProfiles(activeTenantId, priorityProfileMris),
    enabled: Boolean(activeTenantId) && priorityProfileMris.length > 0,
    staleTime: 3_600_000,
    gcTime: PROFILE_QUERY_GC_MS,
    placeholderData: (previousData) => previousData,
    retry: 2,
    retryDelay: (attempt) => Math.min(2000 * 2 ** attempt, 12_000),
  });

  const backgroundAvatarQuery = useQuery({
    queryKey: teamsKeys.profileAvatars(activeTenantId, profileMriSignature),
    queryFn: () => fetchProfiles(activeTenantId, profileMris),
    enabled:
      Boolean(activeTenantId) &&
      profileMris.length > 0 &&
      profileMriSignature !== prioritySignature &&
      priorityAvatarQuery.isSuccess,
    staleTime: 3_600_000,
    gcTime: PROFILE_QUERY_GC_MS,
    placeholderData: () => priorityAvatarQuery.data,
    retry: 2,
    retryDelay: (attempt) => Math.min(2000 * 2 ** attempt, 12_000),
  });

  const profilePresentationData =
    backgroundAvatarQuery.data ?? priorityAvatarQuery.data;

  return useMemo(
    () => normalizeProfilePresentation(profilePresentationData),
    [profilePresentationData],
  );
}

async function fetchPresence(
  tenantId: string | undefined,
  mris: string[],
): Promise<Record<string, PresenceInfo>> {
  const cachedPresence = await getCachedPresence(mris);
  const presence: Record<string, PresenceInfo> = Object.fromEntries(
    Object.entries(cachedPresence).map(([mri, info]) => [
      canonAvatarMri(mri),
      info,
    ]),
  );
  const missingMris = mris.filter((mri) => !(canonAvatarMri(mri) in presence));
  if (missingMris.length === 0) {
    return presence;
  }

  const client = await getOrCreateClient(tenantId);

  for (
    let index = 0;
    index < missingMris.length;
    index += PRESENCE_BATCH_SIZE
  ) {
    const batch = missingMris.slice(index, index + PRESENCE_BATCH_SIZE);
    try {
      const batchPresence = await client.getPresence(batch);
      for (const [mri, info] of Object.entries(batchPresence)) {
        presence[canonAvatarMri(mri)] = info;
      }
    } catch (error) {
      if (Object.keys(presence).length > 0) {
        continue;
      }
      throw error;
    }
  }

  return presence;
}

export function useTeamsPresence(args: {
  conversations: Conversation[];
  messages?: Message[];
  selfSkypeId?: string;
}) {
  const { activeTenantId } = useTeamsAccountContext();
  const documentVisible = useDocumentVisibility();
  const resumeReady = useResumeCooldown();
  const deferredConversations = useDeferredValue(args.conversations);
  const deferredMessages = useDeferredValue(args.messages ?? EMPTY_MESSAGES);
  const presenceMris = useMemo(
    () =>
      uniqueMris(
        collectProfileAvatarMris({
          conversations: deferredConversations,
          messages: deferredMessages,
          selfSkypeId: args.selfSkypeId,
        }),
      ),
    [deferredConversations, deferredMessages, args.selfSkypeId],
  );
  const signature = useMemo(
    () => [...presenceMris].sort().join("\x1f"),
    [presenceMris],
  );

  const query = useQuery({
    queryKey: teamsKeys.presence(activeTenantId, signature),
    queryFn: () => fetchPresence(activeTenantId, presenceMris),
    enabled: documentVisible && resumeReady && presenceMris.length > 0,
    staleTime: 30_000,
    gcTime: 5 * 60_000,
    placeholderData: (previousData) => previousData,
    refetchInterval: () => (documentVisible && resumeReady ? 60_000 : false),
    refetchIntervalInBackground: false,
    refetchOnWindowFocus: false,
    refetchOnReconnect: false,
    retry: 1,
  });

  return query.data ?? EMPTY_PRESENCE_BY_MRI;
}

export function useMaintainTeamsAvailability(enabled: boolean) {
  const { activeTenantId } = useTeamsAccountContext();
  const documentVisible = useDocumentVisibility();
  const resumeReady = useResumeCooldown();

  useQuery({
    queryKey: teamsKeys.selfAvailability(activeTenantId),
    queryFn: async () => {
      const client = await getOrCreateClient(activeTenantId);
      await client.setAvailability("Available");
      return Date.now();
    },
    enabled: enabled && documentVisible && resumeReady,
    staleTime: SELF_AVAILABILITY_HEARTBEAT_MS - 5_000,
    gcTime: 5 * 60_000,
    refetchInterval: SELF_AVAILABILITY_HEARTBEAT_MS,
    refetchIntervalInBackground: false,
    refetchOnWindowFocus: false,
    retry: 1,
  });
}
