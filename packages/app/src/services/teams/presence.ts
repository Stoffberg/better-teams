import type { PresenceInfo } from "@better-teams/core/teams/types";
import { getTeamsClient } from "./client";

export const teamsPresenceService = {
  async getPresence(
    tenantId: string | undefined,
    mris: string[],
  ): Promise<Record<string, PresenceInfo>> {
    const client = await getTeamsClient(tenantId);
    return client.getPresence(mris);
  },

  async setSelfAvailability(tenantId: string | undefined): Promise<void> {
    const client = await getTeamsClient(tenantId);
    await client.setAvailability("Available");
  },
};
