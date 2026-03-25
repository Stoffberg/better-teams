import { useQuery, useQueryClient } from "@tanstack/react-query";
import {
  createContext,
  type ReactNode,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
} from "react";
import { SqliteWorkspaceShellStore } from "@/lib/sqlite-cache";
import {
  clearClientForTenant,
  getOrCreateClient,
} from "@/lib/teams-client-factory";
import { teamsKeys } from "@/lib/teams-query-keys";
import { TeamsApiClient } from "@/services/teams/api-client";
import type {
  TeamsAccountOption,
  TeamsSessionInfo,
  TeamsWorkspaceShellSnapshot,
} from "@/services/teams/types";

const PREFERRED_TENANT_STORAGE_KEY = "better-teams-preferred-tenant-id";

export type TeamsAccountContextValue = {
  accounts: TeamsAccountOption[];
  activeTenantId?: string;
  selectedTenantId?: string | null;
  pendingTenantId?: string | null;
  isSwitchingAccount: boolean;
  activeSession?: TeamsSessionInfo;
  workspaceShell: TeamsWorkspaceShellSnapshot | null;
  switchAccount: (tenantId: string | null) => void;
  persistedPreference: string | null;
};

const TeamsAccountContext = createContext<TeamsAccountContextValue | null>(
  null,
);

function readPreferredTenantId(): string | null {
  try {
    return localStorage.getItem(PREFERRED_TENANT_STORAGE_KEY);
  } catch {
    return null;
  }
}

function writePreferredTenantId(tenantId: string | null): void {
  try {
    if (tenantId) {
      localStorage.setItem(PREFERRED_TENANT_STORAGE_KEY, tenantId);
      return;
    }
    localStorage.removeItem(PREFERRED_TENANT_STORAGE_KEY);
  } catch {
    // Ignore storage failures in restricted environments.
  }
}

function normalizeAccounts(
  accounts: TeamsAccountOption[] | undefined,
): TeamsAccountOption[] {
  return [...(accounts ?? [])].sort((a, b) =>
    (a.upn ?? "").localeCompare(b.upn ?? ""),
  );
}

function resolveSelectedTenantId(
  accounts: TeamsAccountOption[],
  preferredTenantId: string | null,
): string | undefined {
  if (
    preferredTenantId &&
    accounts.some((account) => account.tenantId === preferredTenantId)
  ) {
    return preferredTenantId;
  }
  return accounts[0]?.tenantId ?? preferredTenantId ?? undefined;
}

function shellSessionForTenant(
  workspaceShell: TeamsWorkspaceShellSnapshot | null,
  tenantId?: string | null,
): TeamsSessionInfo | undefined {
  if (!tenantId) return undefined;
  const tenantShell =
    workspaceShell?.tenants[tenantId] ?? workspaceShell?.tenants.__default__;
  return tenantShell?.session;
}

async function initializeTeamsSession(
  tenantId?: string | null,
): Promise<TeamsSessionInfo> {
  const client = await getOrCreateClient(tenantId ?? undefined);
  const a = client.account;
  return {
    upn: a.upn,
    tenantId: a.tenantId ?? "__default__",
    skypeId: a.skypeId,
    expiresAt: a.expiresAt?.toISOString() ?? null,
    region: a.region,
  };
}

export function TeamsAccountProvider({ children }: { children: ReactNode }) {
  const queryClient = useQueryClient();
  const workspaceShellQuery = useQuery({
    queryKey: ["workspace-shell"],
    queryFn: () => SqliteWorkspaceShellStore.getSnapshot(),
    staleTime: Number.POSITIVE_INFINITY,
    gcTime: Number.POSITIVE_INFINITY,
  });
  const workspaceShell =
    (workspaceShellQuery.data as TeamsWorkspaceShellSnapshot | null) ?? null;

  const [persistedPreference, setPersistedPreference] = useState<string | null>(
    () => readPreferredTenantId(),
  );
  const [pendingTenantId, setPendingTenantId] = useState<string | null>(null);

  const accountsQuery = useQuery({
    queryKey: teamsKeys.accounts(),
    queryFn: async () =>
      normalizeAccounts((await TeamsApiClient.getAvailableAccounts()) ?? []),
    initialData: () => {
      const shellAccounts = normalizeAccounts(workspaceShell?.accounts);
      return shellAccounts.length > 0 ? shellAccounts : undefined;
    },
    staleTime: 30_000,
    gcTime: Number.POSITIVE_INFINITY,
  });

  const accounts = useMemo(
    () => normalizeAccounts(accountsQuery.data ?? workspaceShell?.accounts),
    [accountsQuery.data, workspaceShell?.accounts],
  );

  useEffect(() => {
    if (accounts.length === 0) return;
    void SqliteWorkspaceShellStore.updateAccounts(accounts);
  }, [accounts]);
  const selectedTenantId = resolveSelectedTenantId(
    accounts,
    persistedPreference,
  );
  const shellSession = shellSessionForTenant(workspaceShell, selectedTenantId);

  const sessionQuery = useQuery({
    queryKey: teamsKeys.session(selectedTenantId),
    queryFn: async () => initializeTeamsSession(selectedTenantId),
    initialData: shellSession,
    initialDataUpdatedAt: () => {
      if (!selectedTenantId) return undefined;
      return workspaceShell?.tenants[selectedTenantId]?.updatedAt;
    },
    enabled: true,
    staleTime: 30_000,
    gcTime: Number.POSITIVE_INFINITY,
  });

  const switchAccount = useCallback(
    (tenantId: string | null) => {
      const nextTenantId = tenantId ?? null;
      if (nextTenantId === selectedTenantId) return;
      const previousTenantId = persistedPreference;
      setPersistedPreference(nextTenantId);
      setPendingTenantId(nextTenantId);
      clearClientForTenant(nextTenantId);
      void queryClient
        .fetchQuery({
          queryKey: teamsKeys.session(nextTenantId),
          queryFn: () => initializeTeamsSession(nextTenantId),
          staleTime: 30_000,
        })
        .then(() => {
          writePreferredTenantId(nextTenantId);
          setPendingTenantId((current) =>
            current === nextTenantId ? null : current,
          );
        })
        .catch(() => {
          setPersistedPreference(previousTenantId);
          clearClientForTenant(nextTenantId);
          setPendingTenantId((current) =>
            current === nextTenantId ? null : current,
          );
        });
    },
    [persistedPreference, queryClient, selectedTenantId],
  );

  const activeTenantId =
    selectedTenantId ?? sessionQuery.data?.tenantId ?? shellSession?.tenantId;

  const value = useMemo<TeamsAccountContextValue>(
    () => ({
      accounts,
      activeTenantId,
      selectedTenantId,
      pendingTenantId,
      isSwitchingAccount:
        pendingTenantId != null && pendingTenantId === selectedTenantId,
      activeSession: sessionQuery.data ?? shellSession,
      workspaceShell,
      switchAccount,
      persistedPreference,
    }),
    [
      accounts,
      activeTenantId,
      pendingTenantId,
      persistedPreference,
      selectedTenantId,
      sessionQuery.data,
      shellSession,
      switchAccount,
      workspaceShell,
    ],
  );

  return (
    <TeamsAccountContext.Provider value={value}>
      {children}
    </TeamsAccountContext.Provider>
  );
}

export function useTeamsAccountContext(): TeamsAccountContextValue {
  const value = useContext(TeamsAccountContext);
  if (!value) {
    throw new Error(
      "useTeamsAccountContext must be used within TeamsAccountProvider",
    );
  }
  return value;
}
