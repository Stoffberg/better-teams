import { createAsyncStoragePersister } from "@tanstack/query-async-storage-persister";
import { QueryClient } from "@tanstack/react-query";
import { PersistQueryClientProvider } from "@tanstack/react-query-persist-client";
import { type ReactNode, useState } from "react";
import { SqliteQueryPersister } from "@/lib/sqlite-cache";
import { teamsKeys } from "@/lib/teams-query-keys";

const ONE_DAY_MS = 24 * 60 * 60 * 1000;

const persister = createAsyncStoragePersister({
  storage: SqliteQueryPersister.getStorage(),
  throttleTime: 15_000,
});

export function shouldPersistQuery(query: {
  queryKey: readonly unknown[];
}): boolean {
  const rootKey = query.queryKey[0];
  return (
    rootKey !== teamsKeys.all[0] && rootKey !== "open-conversation-request"
  );
}

export function QueryProvider({ children }: { children: ReactNode }) {
  const [client] = useState(
    () =>
      new QueryClient({
        defaultOptions: {
          queries: {
            staleTime: 60_000,
            gcTime: ONE_DAY_MS,
            retry: false,
            refetchOnWindowFocus: false,
            refetchOnReconnect: false,
          },
        },
      }),
  );

  return (
    <PersistQueryClientProvider
      client={client}
      persistOptions={{
        persister,
        maxAge: ONE_DAY_MS,
        buster: "v4-no-teams-query-persist",
        dehydrateOptions: {
          shouldDehydrateQuery: shouldPersistQuery,
        },
      }}
    >
      {children}
    </PersistQueryClientProvider>
  );
}
