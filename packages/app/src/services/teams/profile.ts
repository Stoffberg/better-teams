import type { TeamsProfilePresentation } from "@better-teams/core/teams/types";
import { getTeamsClient } from "./client";

export const teamsProfileService = {
  async fetchPresentation(
    tenantId: string | undefined,
    mris: string[],
  ): Promise<TeamsProfilePresentation> {
    const client = await getTeamsClient(tenantId);
    return client.fetchProfileAvatarDataUrls(mris);
  },

  async fetchAuthenticatedImageSrc(
    tenantId: string,
    imageUrl: string,
  ): Promise<string | null> {
    const client = await getTeamsClient(tenantId);
    return client.fetchAuthenticatedImageSrc(imageUrl);
  },
};
