/**
 * Manages tenant-scoped TeamsApiClient instances in the webview.
 *
 * Replaces the `tenantClients` Map and `getTenantClient()` pattern
 * from the old Electron `teams-ipc.ts` handler.
 *
 * Uses `@tauri-apps/plugin-http`'s fetch for CORS-free HTTP requests
 * and the SQLite image cache for avatar caching.
 */

import { fetch } from "@tauri-apps/plugin-http";
import { SqliteImageCache } from "@/lib/sqlite-cache";
import {
  TeamsApiClient,
  type TeamsApiDiagnosticEvent,
} from "@/services/teams/api-client";

const tenantClients = new Map<string, TeamsApiClient>();
const tenantClientInitializations = new Map<string, Promise<TeamsApiClient>>();

function tenantKey(tenantId?: string): string {
  return tenantId ?? "__default__";
}

function shouldSuppressFetchErrorLog(
  url: string,
  method: string,
  status: number,
): boolean {
  if (status !== 401) return false;
  if (
    method === "POST" &&
    url === "https://presence.teams.microsoft.com/v1/presence/getpresence"
  ) {
    return true;
  }
  if (
    method === "PUT" &&
    url === "https://presence.teams.microsoft.com/v1/me/forceavailability/"
  ) {
    return true;
  }
  return /^https:\/\/[^/]*asm\.skype\.com\/v1\/objects\/.+\/views\//i.test(url);
}

function createClient(tenantId?: string): TeamsApiClient {
  const onDiagnostic = (event: TeamsApiDiagnosticEvent) => {
    if (event.level === "error") {
      console.error("[teams]", event.message, event.data ?? {});
    } else {
      console.warn("[teams]", event.message, event.data ?? {});
    }
  };

  const TEAMS_ORIGIN = "https://teams.microsoft.com";

  const loggingFetch: typeof fetch = async (input, init) => {
    const url =
      typeof input === "string"
        ? input
        : input instanceof URL
          ? input.href
          : input instanceof Request
            ? input.url
            : String(input);

    // Inject Origin + Referer so Teams CORS checks pass
    const headers = new Headers(init?.headers);
    if (!headers.has("Origin")) headers.set("Origin", TEAMS_ORIGIN);
    if (!headers.has("Referer")) headers.set("Referer", `${TEAMS_ORIGIN}/`);

    try {
      const res = await fetch(input, { ...init, headers });
      if (!res.ok) {
        if (
          shouldSuppressFetchErrorLog(url, init?.method ?? "GET", res.status)
        ) {
          return res;
        }
        const body = await res
          .clone()
          .text()
          .catch(() => "");
        console.error(
          `[fetch] ${init?.method ?? "GET"} ${url} → ${res.status}`,
          body.slice(0, 300),
        );
      }
      return res;
    } catch (err) {
      console.error(`[fetch] ${init?.method ?? "GET"} ${url} THREW:`, err);
      throw err;
    }
  };

  return new TeamsApiClient(tenantId, {
    fetchImpl: loggingFetch,
    getCachedImagePath: (url) => SqliteImageCache.get(url),
    setCachedImagePath: (url, filePath) => SqliteImageCache.set(url, filePath),
    onDiagnostic,
  });
}

/**
 * Get an existing client for the tenant or create + initialize a new one.
 */
export async function getOrCreateClient(
  tenantId?: string,
): Promise<TeamsApiClient> {
  const key = tenantKey(tenantId);
  const existing = tenantClients.get(key);
  if (existing) return existing;

  const initializing = tenantClientInitializations.get(key);
  if (initializing) return initializing;

  const client = createClient(tenantId);
  const initialization = client
    .initialize()
    .then(() => {
      tenantClients.set(key, client);
      return client;
    })
    .finally(() => {
      tenantClientInitializations.delete(key);
    });
  tenantClientInitializations.set(key, initialization);
  return initialization;
}

/**
 * Get a client without initializing it (for cases where you need
 * to call initialize yourself, e.g. to return the session info).
 */
export function getOrCreateUninitializedClient(
  tenantId?: string,
): TeamsApiClient {
  const key = tenantKey(tenantId);
  const existing = tenantClients.get(key);
  if (existing) return existing;

  const client = createClient(tenantId);
  tenantClients.set(key, client);
  return client;
}

/**
 * Clear the client cache (e.g., on logout or account switch).
 */
export function clearClientCache(): void {
  tenantClients.clear();
  tenantClientInitializations.clear();
}

export function clearClientForTenant(tenantId?: string | null): void {
  const key = tenantKey(tenantId ?? undefined);
  tenantClients.delete(key);
  tenantClientInitializations.delete(key);
}
