import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { render, screen } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { ChatWorkspace } from "@/components/chat/ChatWorkspace";
import { TeamsAccountProvider } from "@/providers/TeamsAccountProvider";
import { ThemeProvider } from "@/providers/ThemeProvider";

// Mock the teams-client-factory and electron-bridge modules
vi.mock("@/lib/teams-client-factory", () => ({
  getOrCreateClient: vi.fn().mockResolvedValue({
    initialize: vi.fn().mockResolvedValue(undefined),
    account: {
      upn: "user@test.com",
      tenantId: "t1",
      skypeId: "s",
      expiresAt: new Date(),
      region: "amer",
    },
    getAllConversations: vi.fn().mockResolvedValue({ conversations: [] }),
    getFavoriteConversations: vi.fn().mockResolvedValue({ conversations: [] }),
    setConversationFavorite: vi.fn().mockResolvedValue(undefined),
    getMessages: vi.fn().mockResolvedValue({ messages: [] }),
    sendMessage: vi.fn().mockResolvedValue(undefined),
    fetchProfileAvatarDataUrls: vi.fn().mockResolvedValue({
      avatars: {},
      displayNames: {},
      emails: {},
      jobTitles: {},
    }),
    getMessagesByUrl: vi.fn().mockResolvedValue({ messages: [] }),
    getAnchoredMessages: vi.fn().mockResolvedValue({ messages: [] }),
  }),
  getOrCreateUninitializedClient: vi.fn(),
  clearClientCache: vi.fn(),
}));

vi.mock("@/lib/electron-bridge", () => ({
  cacheImageFile: vi.fn(),
  extractTokens: vi.fn().mockResolvedValue([
    {
      host: ".teams.microsoft.com",
      name: "authtoken",
      token: "fake",
      upn: "user@test.com",
      tenantId: "t1",
      skypeId: "s",
      expiresAt: new Date(Date.now() + 3600000),
    },
  ]),
  filePathToAssetUrl: vi.fn((filePath: string) => `asset://${filePath}`),
  getAuthToken: vi.fn().mockResolvedValue({
    host: ".teams.microsoft.com",
    name: "authtoken",
    token: "fake",
    upn: "user@test.com",
    tenantId: "t1",
    skypeId: "s",
    expiresAt: new Date(Date.now() + 3600000),
  }),
  getAvailableAccounts: vi.fn().mockResolvedValue([]),
  removeCachedImageFiles: vi.fn().mockResolvedValue(undefined),
}));

vi.mock("@/lib/sqlite-cache", () => ({
  SqliteWorkspaceShellStore: {
    getSnapshot: vi.fn().mockResolvedValue(null),
    updateAccounts: vi.fn().mockResolvedValue(undefined),
    updateSession: vi.fn().mockResolvedValue(undefined),
    updateConversations: vi.fn().mockResolvedValue(undefined),
  },
  SqliteProfileCache: {
    lookupProfiles: vi.fn().mockResolvedValue({
      presentation: {
        avatars: {},
        displayNames: {},
        emails: {},
        jobTitles: {},
      },
      missingMris: [],
    }),
    storeProfiles: vi.fn().mockResolvedValue(undefined),
    merge: vi.fn().mockImplementation((a: unknown, b: unknown) => ({
      ...(a as Record<string, unknown>),
      ...(b as Record<string, unknown>),
    })),
  },
  SqliteImageCache: {
    get: vi.fn().mockResolvedValue(null),
    set: vi.fn().mockResolvedValue(undefined),
  },
  SqliteQueryPersister: {
    getStorage: vi.fn().mockReturnValue({
      getItem: vi.fn().mockResolvedValue(null),
      setItem: vi.fn().mockResolvedValue(undefined),
      removeItem: vi.fn().mockResolvedValue(undefined),
    }),
  },
}));

function renderApp() {
  const client = new QueryClient({
    defaultOptions: {
      queries: { retry: false, gcTime: 0, refetchOnWindowFocus: false },
    },
  });
  return render(
    <ThemeProvider>
      <QueryClientProvider client={client}>
        <TeamsAccountProvider>
          <div className="flex h-screen flex-col overflow-hidden">
            <ChatWorkspace />
          </div>
        </TeamsAccountProvider>
      </QueryClientProvider>
    </ThemeProvider>,
  );
}

describe("App", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    localStorage.clear();
  });

  it("renders chat workspace", async () => {
    renderApp();
    expect(
      await screen.findByRole("searchbox", { name: /search/i }),
    ).toBeInTheDocument();
  });
});
