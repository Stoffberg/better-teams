import { ChatWorkspace } from "@better-teams/app/features/chat/workspace/ChatWorkspace";
import { TeamsAccountProvider } from "@better-teams/app/providers/TeamsAccountProvider";
import { ThemeProvider } from "@better-teams/app/providers/ThemeProvider";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { render, screen } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("@better-teams/core/teams/client/factory", () => ({
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
  clearClientCache: vi.fn(),
}));

vi.mock("@better-teams/app/services/desktop/runtime", () => ({
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
  getCachedImageFile: vi.fn().mockResolvedValue(null),
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
  hasCachedImageFile: vi.fn().mockResolvedValue(false),
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
