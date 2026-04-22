import { getTeamsRuntime } from "./runtime";
import {
  TeamsApiClient,
  type TeamsApiDiagnosticEvent,
} from "./teams/api-client";

const tenantClients = new Map<string, TeamsApiClient>();
const tenantClientInitializations = new Map<string, Promise<TeamsApiClient>>();

function tenantKey(tenantId?: string): string {
  return tenantId ?? "__default__";
}

function shouldSuppressFetchErrorLog(
  url: string,
  method: string,
  status: number,
  headers?: Headers,
): boolean {
  if (headers?.get("Accept")?.includes("image/")) return true;
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

  const runtime = getTeamsRuntime();
  const loggingFetch: typeof runtime.fetch = async (input, init) => {
    const url =
      typeof input === "string"
        ? input
        : input instanceof URL
          ? input.href
          : input instanceof Request
            ? input.url
            : String(input);

    const headers = new Headers(init?.headers);
    if (!headers.has("Origin")) headers.set("Origin", TEAMS_ORIGIN);
    if (!headers.has("Referer")) headers.set("Referer", `${TEAMS_ORIGIN}/`);

    try {
      const res = await runtime.fetch(input, { ...init, headers });
      if (!res.ok) {
        if (
          shouldSuppressFetchErrorLog(
            url,
            init?.method ?? "GET",
            res.status,
            headers,
          )
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
      if (headers.get("Accept")?.includes("image/")) {
        throw err;
      }
      console.error(`[fetch] ${init?.method ?? "GET"} ${url} THREW:`, err);
      throw err;
    }
  };

  return new TeamsApiClient(tenantId, {
    runtime,
    fetchImpl: loggingFetch,
    onDiagnostic,
  });
}

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

export function clearClientCache(): void {
  tenantClients.clear();
  tenantClientInitializations.clear();
}

export function clearClientForTenant(tenantId?: string | null): void {
  const key = tenantKey(tenantId ?? undefined);
  tenantClients.delete(key);
  tenantClientInitializations.delete(key);
}
