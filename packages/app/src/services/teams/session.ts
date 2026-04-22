import { clearClientForTenant } from "@better-teams/core/teams/client/factory";
import type { TeamsSessionInfo } from "@better-teams/core/teams/types";
import { getTeamsClient } from "./client";

export const teamsSessionService = {
  async initialize(tenantId?: string | null): Promise<TeamsSessionInfo> {
    const client = await getTeamsClient(tenantId);
    const account = client.account;
    return {
      upn: account.upn,
      tenantId: account.tenantId ?? "__default__",
      skypeId: account.skypeId,
      expiresAt: account.expiresAt?.toISOString() ?? null,
      region: account.region,
    };
  },

  clearTenantClient(tenantId?: string | null): void {
    clearClientForTenant(tenantId ?? null);
  },
};
