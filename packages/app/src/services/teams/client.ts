import { getOrCreateClient } from "@better-teams/core/teams/client/factory";

export async function getTeamsClient(tenantId?: string | null) {
  return getOrCreateClient(tenantId ?? undefined);
}
