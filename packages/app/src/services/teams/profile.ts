import { getCachedProfilePresentation } from "@better-teams/app/services/desktop/runtime";
import type { TeamsProfilePresentation } from "@better-teams/core/teams/types";
import { getTeamsClient } from "./client";

export const teamsProfileService = {
  async fetchPresentation(
    tenantId: string | undefined,
    mris: string[],
  ): Promise<TeamsProfilePresentation> {
    try {
      const client = await getTeamsClient(tenantId);
      return await client.fetchProfileAvatarDataUrls(mris);
    } catch (error) {
      const cached = await getCachedProfilePresentation(mris);
      if (!cached) throw error;
      const hasCachedData =
        Object.keys(cached.avatarThumbs).length > 0 ||
        Object.keys(cached.avatarFull).length > 0 ||
        Object.keys(cached.displayNames).length > 0;
      if (hasCachedData) return cached;
      throw error;
    }
  },

  async fetchAuthenticatedImageSrc(
    tenantId: string,
    imageUrl: string,
  ): Promise<string | null> {
    const client = await getTeamsClient(tenantId);
    return client.fetchAuthenticatedImageSrc(imageUrl);
  },
};
