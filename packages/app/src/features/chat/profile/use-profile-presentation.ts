import { useTeamsAccountContext } from "@better-teams/app/providers/TeamsAccountProvider";
import { teamsProfileService } from "@better-teams/app/services/teams/profile";
import { teamsKeys } from "@better-teams/app/services/teams/query-keys";
import { collectProfileAvatarMris } from "@better-teams/core/teams/profile/avatars";
import { TeamsProfilePresentationSchema } from "@better-teams/core/teams/schemas";
import type {
  Conversation,
  Message,
  TeamsProfilePresentation,
} from "@better-teams/core/teams/types";
import { useQuery } from "@tanstack/react-query";
import { useDeferredValue, useMemo } from "react";

const PRIORITY_AVATAR_CONVERSATIONS = 12;
const PROFILE_QUERY_GC_MS = 5 * 60_000;
const EMPTY_MESSAGES: Message[] = [];

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

async function fetchProfiles(
  tenantId: string | undefined,
  mris: string[],
): Promise<TeamsProfilePresentation> {
  if (mris.length === 0) return normalizeProfilePresentation(null);
  return teamsProfileService.fetchPresentation(tenantId, mris);
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
