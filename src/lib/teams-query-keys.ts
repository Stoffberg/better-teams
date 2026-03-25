export const teamsKeys = {
  all: ["teams"] as const,
  scope: (tenantId?: string | null) => tenantId ?? "__default__",
  accounts: () => [...teamsKeys.all, "accounts"] as const,
  session: (tenantId?: string | null) =>
    [...teamsKeys.all, "session", teamsKeys.scope(tenantId)] as const,
  conversations: (tenantId?: string | null) =>
    [...teamsKeys.all, "conversations", teamsKeys.scope(tenantId)] as const,
  thread: (tenantId: string | null | undefined, conversationId: string) =>
    [
      ...teamsKeys.all,
      "thread",
      teamsKeys.scope(tenantId),
      conversationId,
    ] as const,
  threadCache: (tenantId: string | null | undefined, conversationId: string) =>
    [
      ...teamsKeys.all,
      "thread-cache",
      teamsKeys.scope(tenantId),
      conversationId,
    ] as const,
  threadMembers: (
    tenantId: string | null | undefined,
    conversationId: string,
  ) =>
    [
      ...teamsKeys.all,
      "thread-members",
      teamsKeys.scope(tenantId),
      conversationId,
    ] as const,
  threadConsumptionHorizons: (
    tenantId: string | null | undefined,
    conversationId: string,
  ) =>
    [
      ...teamsKeys.all,
      "thread-consumption-horizons",
      teamsKeys.scope(tenantId),
      conversationId,
    ] as const,
  profileAvatars: (tenantId: string | null | undefined, mriSignature: string) =>
    [
      ...teamsKeys.all,
      "profileAvatars",
      teamsKeys.scope(tenantId),
      mriSignature,
    ] as const,
  presence: (tenantId: string | null | undefined, mriSignature: string) =>
    [
      ...teamsKeys.all,
      "presence",
      teamsKeys.scope(tenantId),
      mriSignature,
    ] as const,
  selfAvailability: (tenantId?: string | null) =>
    [...teamsKeys.all, "selfAvailability", teamsKeys.scope(tenantId)] as const,
};
