import { useQuery } from "@tanstack/react-query";
import { getOrCreateClient } from "./teams-client-factory";

/**
 * React hook that fetches an authenticated image from Teams AMS/CDN.
 * Returns a local asset URL (or null while loading / on failure).
 */
export function useAuthenticatedImage(
  imageUrl: string | undefined,
  tenantId: string | undefined | null,
): { src: string | null; loading: boolean; error: boolean } {
  const queryImageUrl = imageUrl ?? null;
  const queryTenantId = tenantId ?? null;
  const query = useQuery({
    queryKey: ["authenticated-image", queryTenantId, queryImageUrl],
    enabled: Boolean(queryImageUrl && queryTenantId),
    staleTime: 30 * 60_000,
    gcTime: 2 * 60 * 60_000,
    retry: false,
    queryFn: async () => {
      if (!queryImageUrl || !queryTenantId) return null;
      const client = await getOrCreateClient(queryTenantId);
      return client.fetchAuthenticatedImageSrc(queryImageUrl);
    },
  });

  return {
    src: query.data ?? null,
    loading: query.isLoading,
    error: query.isError || (query.isFetched && !query.data),
  };
}
