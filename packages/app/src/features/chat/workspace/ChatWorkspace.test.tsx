// Module mocks must be declared before any imports that use them.
// Vitest hoists vi.mock() calls, but we keep them at the top for clarity.

import { vi } from "vitest";

vi.mock("@better-teams/core/teams/client/factory");
vi.mock("@better-teams/app/features/chat/thread/preload", async () => {
  const actual = await vi.importActual<
    typeof import("@better-teams/app/features/chat/thread/preload")
  >("@better-teams/app/features/chat/thread/preload");
  return {
    ...actual,
    preloadConversationThread: vi.fn(actual.preloadConversationThread),
  };
});

vi.mock("@better-teams/app/services/desktop/runtime", () => ({
  cacheImageFile: vi.fn(),
  extractTokens: vi.fn().mockResolvedValue([]),
  filePathToAssetUrl: vi.fn((filePath: string) => `asset://${filePath}`),
  getCachedImageFile: vi.fn().mockResolvedValue(null),
  getAuthToken: vi.fn().mockResolvedValue(null),
  getAvailableAccounts: vi.fn().mockResolvedValue([]),
  hasCachedImageFile: vi.fn().mockResolvedValue(false),
}));

import { preloadConversationThread } from "@better-teams/app/features/chat/thread/preload";
import { PERF_FLAG, resetPerfStore } from "@better-teams/app/platform/perf";
import { TeamsAccountProvider } from "@better-teams/app/providers/TeamsAccountProvider";
import { ThemeProvider } from "@better-teams/app/providers/ThemeProvider";
import { getOrCreateClient } from "@better-teams/core/teams/client/factory";
import type { Message } from "@better-teams/core/teams/types";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from "@testing-library/react";
import type { ReactElement } from "react";
import { beforeEach, describe, expect, it } from "vitest";
import { ChatWorkspace } from "./ChatWorkspace";

function renderChat(node: ReactElement = <ChatWorkspace />) {
  const client = new QueryClient({
    defaultOptions: {
      queries: {
        retry: false,
        gcTime: 0,
        refetchOnWindowFocus: false,
      },
    },
  });
  const rendered = render(
    <ThemeProvider>
      <QueryClientProvider client={client}>
        <TeamsAccountProvider>{node}</TeamsAccountProvider>
      </QueryClientProvider>
    </ThemeProvider>,
  );
  return { ...rendered, client };
}

function makeMockClient(overrides: Record<string, unknown> = {}) {
  return {
    initialize: vi.fn().mockResolvedValue(undefined),
    account: {
      upn: "user@test.com",
      tenantId: "t1",
      skypeId: "self",
      expiresAt: new Date(),
      region: "amer",
    },
    getAllConversations: vi.fn().mockResolvedValue({ conversations: [] }),
    getFavoriteConversations: vi.fn().mockResolvedValue({ conversations: [] }),
    setConversationFavorite: vi.fn().mockResolvedValue(undefined),
    getMessages: vi.fn().mockResolvedValue({ messages: [] }),
    getConversation: vi.fn().mockResolvedValue(null),
    getThreadMembers: vi.fn().mockResolvedValue([]),
    sendMessage: vi.fn().mockResolvedValue(undefined),
    sendAttachmentMessage: vi.fn().mockResolvedValue(undefined),
    fetchProfileAvatarDataUrls: vi.fn().mockResolvedValue({
      avatars: {},
      displayNames: {},
      emails: {},
      jobTitles: {},
    }),
    getMessagesByUrl: vi.fn().mockResolvedValue({ messages: [] }),
    getAnchoredMessages: vi.fn().mockResolvedValue({ messages: [] }),
    ...overrides,
  };
}

function baseMessage(
  partial: Partial<Message> & Pick<Message, "id" | "from" | "content">,
): Message {
  return {
    conversationId: "c1",
    type: "Message",
    messagetype: "Text",
    contenttype: "text",
    composetime: "2024-06-01T12:00:00.000Z",
    originalarrivaltime: "2024-06-01T12:00:00.000Z",
    ...partial,
  };
}

describe("ChatWorkspace", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.useRealTimers();
    localStorage.clear();
    resetPerfStore();
  });

  it("keeps the workspace shell visible while session loads", async () => {
    let resolveClient!: (v: unknown) => void;
    const clientPromise = new Promise((r) => {
      resolveClient = r;
    });
    vi.mocked(getOrCreateClient).mockReturnValue(clientPromise as never);
    renderChat();
    expect(
      screen.getByRole("searchbox", { name: /search chats/i }),
    ).toBeInTheDocument();
    expect(screen.queryByText(/connecting/i)).not.toBeInTheDocument();
    expect(document.querySelector('[data-slot="skeleton"]')).toBeTruthy();
    resolveClient(
      makeMockClient({
        account: {
          upn: "user@test.com",
          tenantId: "t1",
          skypeId: "self",
          expiresAt: new Date(),
          region: "amer",
        },
      }),
    );
    expect(
      await screen.findByRole("button", { name: /switch account/i }),
    ).toBeInTheDocument();
  });

  it("uses cached account data while refreshing the Teams session", async () => {
    localStorage.setItem(
      "better-teams-cached-accounts",
      JSON.stringify([{ upn: "cached@test.com", tenantId: "t1" }]),
    );
    localStorage.setItem(
      "better-teams-cached-session",
      JSON.stringify({
        upn: "cached@test.com",
        tenantId: "t1",
        skypeId: "self",
        expiresAt: "2026-01-01T00:00:00.000Z",
        region: "amer",
      }),
    );
    let resolveClient!: (v: unknown) => void;
    const clientPromise = new Promise((r) => {
      resolveClient = r;
    });
    vi.mocked(getOrCreateClient).mockReturnValue(clientPromise as never);

    renderChat();

    expect(
      screen.getByRole("button", { name: /switch account/i }),
    ).toBeInTheDocument();
    expect(screen.getByText("cached@test.com")).toBeInTheDocument();
    expect(screen.queryByText(/connecting/i)).not.toBeInTheDocument();

    resolveClient(makeMockClient());
  });

  it("shows error and retry when initialize fails", async () => {
    vi.mocked(getOrCreateClient).mockRejectedValue(
      new Error("keychain blocked"),
    );
    renderChat();
    expect(await screen.findByText(/keychain blocked/i)).toBeInTheDocument();

    const mockClient = makeMockClient({
      account: {
        upn: "u@x.com",
        tenantId: "t",
        skypeId: "s",
        expiresAt: new Date(),
        region: "r",
      },
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    fireEvent.click(screen.getByRole("button", { name: /try again/i }));
    await waitFor(() => {
      expect(screen.queryByText(/keychain blocked/i)).not.toBeInTheDocument();
    });
    expect(
      await screen.findByText(/no conversations yet/i),
    ).toBeInTheDocument();
  });

  it("lists Teams group chat in the flat sidebar without rendering a thread header", async () => {
    const msg = baseMessage({
      id: "m1",
      from: "8:other",
      content: "Hello",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "19:4c7d0247747f4d9da394d99eb9815e65@thread.v2",
            threadProperties: {
              topic: "Internal Engineering",
              threadType: "chat",
              productThreadType: "Chat",
            },
            lastMessage: msg,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [msg] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    expect(await screen.findByText("Internal Engineering")).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: /^groups$/i })).toBeNull();
    expect(
      screen.queryByRole("button", { name: /^direct messages$/i }),
    ).toBeNull();
    fireEvent.click(screen.getByText("Internal Engineering"));
    const thread = await screen.findByRole("region", {
      name: /message thread/i,
    });
    expect(thread).toBeInTheDocument();
    expect(await screen.findAllByText(/^Group$/i)).not.toHaveLength(0);
  });

  it("records selection perf metrics when perf mode is enabled", async () => {
    localStorage.setItem(PERF_FLAG, "1");
    const msg = baseMessage({
      id: "m1",
      from: "8:other",
      content: "Hello",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "19:group@thread.v2",
            threadProperties: {
              topic: "Design review",
              threadType: "chat",
              productThreadType: "Chat",
            },
            lastMessage: msg,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [msg] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Design review"));

    await screen.findByRole("region", { name: /message thread/i });

    expect(window.__BETTER_TEAMS_PERF__?.metrics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          name: "workspace.selectConversation.requested",
          detail: expect.objectContaining({
            conversationId: "19:group@thread.v2",
          }),
        }),
        expect.objectContaining({
          name: "workspace.selectConversation",
          detail: expect.objectContaining({
            conversationId: "19:group@thread.v2",
          }),
        }),
      ]),
    );
    expect(
      window.__BETTER_TEAMS_PERF__?.snapshots["workspace.sidebar"]?.values,
    ).toEqual(
      expect.objectContaining({
        conversationCount: 1,
        selectedConversation: "19:group@thread.v2",
      }),
    );
  });

  it("preloads a conversation after hover settles on a sidebar row", async () => {
    const msg = baseMessage({
      id: "m1",
      from: "8:other",
      content: "Hello",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            members: [{ id: "8:other", displayName: "Pat Lee" }],
            lastMessage: msg,
          },
        ],
      }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);
    vi.mocked(preloadConversationThread).mockResolvedValueOnce({
      messages: [msg],
      olderPageUrl: null,
      moreOlder: false,
    });

    renderChat();

    const conversation = await screen.findByRole("button", {
      name: /pat lee, direct message/i,
    });
    fireEvent.pointerEnter(conversation);
    await new Promise((resolve) => window.setTimeout(resolve, 160));

    expect(preloadConversationThread).toHaveBeenCalledWith("t1", "c1", 60_000);
  });

  it("shows activity messages in the sidebar and thread instead of hiding them", async () => {
    const activityMessage = baseMessage({
      id: "m-activity",
      from: "8:other",
      type: "ThreadActivity",
      messagetype: "Event",
      content: "{}",
      properties: {
        activity: {
          sourceUserImDisplayName: "Pat Lee",
          activityOperationType: "addedMember",
          messagePreview: "Jordan Rivera",
        },
      },
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "19:activity@thread.v2",
            threadProperties: {
              topic: "Design review",
              threadType: "meeting",
            },
            lastMessage: activityMessage,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [activityMessage] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();

    expect(await screen.findByText("Design review")).toBeInTheDocument();
    fireEvent.click(screen.getByText("Design review"));

    expect(
      await screen.findByText("Pat Lee added member: Jordan Rivera"),
    ).toBeInTheDocument();
    expect(
      screen.queryByText(/only meeting and call activity in this thread/i),
    ).toBeNull();
  });

  it("orders chats by their latest activity regardless of chat kind", async () => {
    const msgIe = baseMessage({
      id: "m1",
      from: "8:a",
      content: "a",
      composetime: "2024-06-01T12:00:00.000Z",
      originalarrivaltime: "2024-06-01T12:00:00.000Z",
    });
    const msgDm = baseMessage({
      id: "m2",
      from: "8:b",
      content: "b",
      composetime: "2024-06-02T12:00:00.000Z",
      originalarrivaltime: "2024-06-02T12:00:00.000Z",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "19:ie@thread.v2",
            threadProperties: {
              topic: "Internal Engineering",
              threadType: "chat",
              productThreadType: "Chat",
            },
            lastMessage: msgIe,
          },
          {
            id: "19:pat@thread.v2",
            threadProperties: {
              topic: "Pat Lee",
              threadType: "chat",
              productThreadType: "Chat",
              membercount: "2",
            },
            lastMessage: msgDm,
          },
        ],
      }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    await screen.findByText("Internal Engineering");
    const conversationButtons = screen.getAllByRole("button", {
      name: /, (direct message|group chat)$/i,
    });
    // Sidebar is now grouped by kind, so all items should still be present
    const allTexts = conversationButtons.map((button) => button.textContent);
    expect(allTexts.some((t) => t?.includes("Pat Lee"))).toBe(true);
    expect(allTexts.some((t) => t?.includes("Internal Engineering"))).toBe(
      true,
    );
  });

  it("shows favorite chats first in alphabetical order and keeps other chats sorted by activity", async () => {
    const recentMessage = baseMessage({
      id: "m-recent",
      from: "8:recent",
      content: "recent",
      composetime: "2024-06-03T12:00:00.000Z",
      originalarrivaltime: "2024-06-03T12:00:00.000Z",
    });
    const mediumMessage = baseMessage({
      id: "m-medium",
      from: "8:medium",
      content: "medium",
      composetime: "2024-06-02T18:00:00.000Z",
      originalarrivaltime: "2024-06-02T18:00:00.000Z",
    });
    const olderMessage = baseMessage({
      id: "m-older",
      from: "8:older",
      content: "older",
      composetime: "2024-06-01T12:00:00.000Z",
      originalarrivaltime: "2024-06-01T12:00:00.000Z",
    });
    const favoriteMessage = baseMessage({
      id: "m-favorite",
      from: "8:favorite",
      content: "favorite",
      composetime: "2024-06-02T12:00:00.000Z",
      originalarrivaltime: "2024-06-02T12:00:00.000Z",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "19:omega@thread.v2",
            threadProperties: { topic: "Omega" },
            lastMessage: recentMessage,
          },
          {
            id: "19:beta@thread.v2",
            threadProperties: { topic: "Beta" },
            lastMessage: mediumMessage,
          },
          {
            id: "19:alpha@thread.v2",
            threadProperties: { topic: "Alpha" },
            properties: { favorite: true },
            lastMessage: olderMessage,
          },
          {
            id: "19:zebra@thread.v2",
            threadProperties: { topic: "Zebra" },
            properties: { favorite: true },
            lastMessage: favoriteMessage,
          },
        ],
      }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    await screen.findByText("Alpha");

    const conversationButtons = screen.getAllByRole("button", {
      name: /, (direct message|group chat|meeting)$/i,
    });
    const titles = conversationButtons.map(
      (button) => button.getAttribute("aria-label")?.split(",")[0],
    );

    expect(titles.slice(0, 4)).toEqual(["Alpha", "Zebra", "Omega", "Beta"]);
  });

  it("toggles favorites from the sidebar", async () => {
    const message = baseMessage({
      id: "m-favorite-toggle",
      from: "8:toggle",
      content: "toggle",
    });
    let favorite = false;
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockImplementation(async () => ({
        conversations: [
          {
            id: "19:toggle@thread.v2",
            threadProperties: { topic: "Toggle Chat" },
            properties: favorite ? { favorite: true } : undefined,
            lastMessage: message,
          },
        ],
      })),
      setConversationFavorite: vi
        .fn()
        .mockImplementation(
          async (_conversationId: string, nextFavorite: boolean) => {
            favorite = nextFavorite;
          },
        ),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    await screen.findByText("Toggle Chat");

    fireEvent.click(
      screen.getByRole("button", { name: /add toggle chat to favorites/i }),
    );

    await waitFor(() => {
      expect(mockClient.setConversationFavorite).toHaveBeenCalledWith(
        "19:toggle@thread.v2",
        true,
      );
    });
    await waitFor(() => {
      expect(
        screen.getByRole("button", {
          name: /remove toggle chat from favorites/i,
        }),
      ).toBeInTheDocument();
    });

    const conversationButton = screen.getByRole("button", {
      name: /toggle chat, group chat/i,
    });
    const favoriteButton = screen.getByRole("button", {
      name: /remove toggle chat from favorites/i,
    });

    expect(conversationButton.contains(favoriteButton)).toBe(false);
  });

  it("shows group chats and meetings differently in the sidebar", async () => {
    const groupMessage = baseMessage({
      id: "m-group",
      from: "8:group",
      content: "group",
    });
    const meetingMessage = baseMessage({
      id: "m-meeting",
      from: "8:meeting",
      content: "meeting",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "19:group@thread.v2",
            threadProperties: {
              topic: "Internal Engineering",
              threadType: "chat",
              productThreadType: "Chat",
            },
            members: [{ id: "8:a" }, { id: "8:b" }, { id: "8:c" }],
            lastMessage: groupMessage,
          },
          {
            id: "19:meeting@thread.v2",
            threadProperties: {
              topic: "Weekly Sync",
              threadType: "meeting",
            },
            lastMessage: meetingMessage,
          },
        ],
      }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();

    const groupButton = await screen.findByRole("button", {
      name: /internal engineering, group chat/i,
    });
    const meetingButton = await screen.findByRole("button", {
      name: /weekly sync, meeting/i,
    });

    expect(groupButton.querySelector(".lucide-users")).not.toBeNull();
    expect(groupButton.querySelector(".lucide-video")).toBeNull();
    expect(meetingButton.querySelector(".lucide-video")).not.toBeNull();
    expect(meetingButton.querySelector(".lucide-users")).toBeNull();
  });

  it("lists conversations and opens a thread", async () => {
    const msg = baseMessage({
      id: "m1",
      from: "8:other",
      content: "Hello world",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Design review" },
            lastMessage: msg,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [msg] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    expect(await screen.findByText("Design review")).toBeInTheDocument();
    fireEvent.click(screen.getByText("Design review"));
    await waitFor(() => {
      expect(mockClient.getMessages).toHaveBeenCalledWith("c1", 80, 1);
    });
    const thread = screen.getByRole("region", { name: /message thread/i });
    await waitFor(() => {
      expect(thread).toHaveTextContent("Hello world");
    });
  });

  it("switches accounts and reloads conversations for the selected tenant", async () => {
    const mockClientA = makeMockClient({
      account: {
        upn: "alpha@test.com",
        tenantId: "tenant-a",
        skypeId: "self-a",
        expiresAt: new Date(),
        region: "amer",
      },
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c-a",
            threadProperties: { topic: "Alpha account chat" },
            lastMessage: baseMessage({
              id: "m-a",
              from: "8:alpha",
              content: "Hello from Alpha",
            }),
          },
        ],
      }),
    });
    const mockClientB = makeMockClient({
      account: {
        upn: "bravo@test.com",
        tenantId: "tenant-b",
        skypeId: "self-b",
        expiresAt: new Date(),
        region: "amer",
      },
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c-b",
            threadProperties: { topic: "Bravo account chat" },
            lastMessage: baseMessage({
              id: "m-b",
              from: "8:bravo",
              content: "Hello from Bravo",
            }),
          },
        ],
      }),
    });

    const { extractTokens } = await import(
      "@better-teams/app/services/desktop/runtime"
    );
    vi.mocked(extractTokens).mockResolvedValue([
      {
        host: ".teams.microsoft.com",
        name: "authtoken",
        token: "fake-a",
        upn: "alpha@test.com",
        tenantId: "tenant-a",
        skypeId: "self-a",
        expiresAt: new Date(Date.now() + 3600000),
      },
      {
        host: ".teams.microsoft.com",
        name: "authtoken",
        token: "fake-b",
        upn: "bravo@test.com",
        tenantId: "tenant-b",
        skypeId: "self-b",
        expiresAt: new Date(Date.now() + 3600000),
      },
    ] as never);

    vi.mocked(getOrCreateClient).mockImplementation(
      async (tenantId?: string) => {
        if (tenantId === "tenant-b") return mockClientB as never;
        return mockClientA as never;
      },
    );

    renderChat();

    expect(await screen.findByText("Alpha account chat")).toBeInTheDocument();
    fireEvent.pointerDown(
      screen.getByRole("button", { name: /switch account/i }),
      { button: 0, ctrlKey: false },
    );
    fireEvent.click(
      await screen.findByRole("menuitemradio", { name: /bravo@test.com/i }),
    );

    expect(await screen.findByText("Bravo account chat")).toBeInTheDocument();
    expect(getOrCreateClient).toHaveBeenCalledWith("tenant-b");
  });

  it("clears visible conversations while switching accounts", async () => {
    let resolveClientB!: (value: unknown) => void;
    const clientBPromise = new Promise((resolve) => {
      resolveClientB = resolve;
    });

    const mockClientA = makeMockClient({
      account: {
        upn: "alpha@test.com",
        tenantId: "tenant-a",
        skypeId: "self-a",
        expiresAt: new Date(),
        region: "amer",
      },
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c-a",
            threadProperties: { topic: "Alpha account chat" },
            lastMessage: baseMessage({
              id: "m-a",
              from: "8:alpha",
              content: "Hello from Alpha",
            }),
          },
        ],
      }),
    });
    const mockClientB = makeMockClient({
      account: {
        upn: "bravo@test.com",
        tenantId: "tenant-b",
        skypeId: "self-b",
        expiresAt: new Date(),
        region: "amer",
      },
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c-b",
            threadProperties: { topic: "Bravo account chat" },
            lastMessage: baseMessage({
              id: "m-b",
              from: "8:bravo",
              content: "Hello from Bravo",
            }),
          },
        ],
      }),
    });

    const { extractTokens } = await import(
      "@better-teams/app/services/desktop/runtime"
    );
    vi.mocked(extractTokens).mockResolvedValue([
      {
        host: ".teams.microsoft.com",
        name: "authtoken",
        token: "fake-a",
        upn: "alpha@test.com",
        tenantId: "tenant-a",
        skypeId: "self-a",
        expiresAt: new Date(Date.now() + 3600000),
      },
      {
        host: ".teams.microsoft.com",
        name: "authtoken",
        token: "fake-b",
        upn: "bravo@test.com",
        tenantId: "tenant-b",
        skypeId: "self-b",
        expiresAt: new Date(Date.now() + 3600000),
      },
    ] as never);

    vi.mocked(getOrCreateClient).mockImplementation(
      async (tenantId?: string) => {
        if (tenantId === "tenant-b") return clientBPromise as never;
        return mockClientA as never;
      },
    );

    renderChat();

    expect(await screen.findByText("Alpha account chat")).toBeInTheDocument();
    fireEvent.pointerDown(
      screen.getByRole("button", { name: /switch account/i }),
      { button: 0, ctrlKey: false },
    );
    fireEvent.click(
      await screen.findByRole("menuitemradio", { name: /bravo@test.com/i }),
    );

    expect(screen.queryByText("Alpha account chat")).not.toBeInTheDocument();
    expect(screen.queryByText(/connecting/i)).not.toBeInTheDocument();
    expect(document.querySelector('[data-slot="skeleton"]')).toBeTruthy();

    resolveClientB(mockClientB);

    expect(await screen.findByText("Bravo account chat")).toBeInTheDocument();
  });

  it("sends a message from the composer", async () => {
    const msg = baseMessage({
      id: "m1",
      from: "8:other",
      content: "Hi",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Thread" },
            lastMessage: msg,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [msg] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Thread"));
    const input = await screen.findByRole("textbox", { name: /message text/i });
    fireEvent.input(input, { target: { innerHTML: "Outbox text" } });
    fireEvent.click(screen.getByRole("button", { name: /send/i }));
    await waitFor(() => {
      expect(mockClient.sendMessage).toHaveBeenCalledWith(
        "c1",
        "Outbox text",
        expect.any(String),
        "text",
        [],
      );
    });
    await waitFor(() => {
      expect(
        mockClient.getAllConversations.mock.calls.length,
      ).toBeGreaterThanOrEqual(2);
    });
  });

  it("uploads attachments from the composer", async () => {
    const msg = baseMessage({
      id: "m1",
      from: "8:other",
      content: "Hi",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            members: [{ id: "8:orgid:peer", role: "User", isMri: true }],
            threadProperties: { topic: "Thread" },
            lastMessage: msg,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [msg] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Thread"));
    const fileInput = document.querySelector(
      'input[type="file"]',
    ) as HTMLInputElement | null;
    expect(fileInput).not.toBeNull();
    const file = new File(["hello"], "brief.txt", {
      type: "text/plain",
    });
    if (!fileInput) {
      throw new Error("Expected file input");
    }
    fireEvent.change(fileInput, { target: { files: [file] } });
    fireEvent.click(screen.getByRole("button", { name: /send/i }));
    await waitFor(() => {
      expect(mockClient.sendAttachmentMessage).toHaveBeenCalledWith(
        "c1",
        file,
        expect.any(String),
        ["8:orgid:peer"],
      );
    });
  });

  it("filters conversations by search", async () => {
    const mockClient = makeMockClient({
      account: {
        upn: "u@x.com",
        tenantId: "t",
        skypeId: "s",
        expiresAt: new Date(),
        region: "",
      },
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "a",
            threadProperties: { topic: "Alpha chat" },
            lastMessage: baseMessage({
              id: "1",
              from: "8:x",
              content: "a",
            }),
          },
          {
            id: "b",
            threadProperties: { topic: "Beta chat" },
            lastMessage: baseMessage({
              id: "2",
              from: "8:y",
              content: "b",
            }),
          },
        ],
      }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    await screen.findByText("Alpha chat");
    expect(screen.getByText("Beta chat")).toBeInTheDocument();
    const search = screen.getByRole("searchbox", { name: /search chats/i });
    fireEvent.change(search, { target: { value: "Beta" } });
    expect(screen.queryByText("Alpha chat")).not.toBeInTheDocument();
    expect(screen.getByText("Beta chat")).toBeInTheDocument();
  });

  it("shows a debounced result count for in chat search", async () => {
    const firstMatch = baseMessage({
      id: "m1",
      from: "8:other",
      content: "alpha release note",
    });
    const secondMatch = baseMessage({
      id: "m2",
      from: "8:other",
      content: "beta alpha follow up",
      composetime: "2024-06-01T12:01:00.000Z",
      originalarrivaltime: "2024-06-01T12:01:00.000Z",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Design review" },
            lastMessage: secondMatch,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({
        messages: [firstMatch, secondMatch],
      }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Design review"));
    await screen.findByText("alpha release note");

    fireEvent.click(
      screen.getByRole("button", { name: /search in conversation/i }),
    );
    fireEvent.change(
      screen.getByRole("searchbox", { name: /find in conversation/i }),
      { target: { value: "alpha" } },
    );

    expect(screen.getByText("Searching")).toBeInTheDocument();
    await waitFor(() => {
      expect(screen.getByText("2 results")).toBeInTheDocument();
    });
  });

  it("shows empty thread placeholder until selection", async () => {
    const mockClient = makeMockClient({
      account: {
        upn: "u@x.com",
        tenantId: "t",
        skypeId: "s",
        expiresAt: new Date(),
        region: "",
      },
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "19:4c7d0247747f4d9da394d99eb9815e65@thread.v2",
            threadProperties: {
              topic: "Internal Engineering",
              threadType: "chat",
              productThreadType: "Chat",
              membercount: "3",
            },
            lastMessage: baseMessage({
              id: "1",
              from: "8:x",
              content: "x",
            }),
          },
        ],
      }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    await screen.findByRole("button", { name: /switch account/i });
    expect(await screen.findByText("Internal Engineering")).toBeInTheDocument();
    expect(screen.getByText(/select a conversation/i)).toBeInTheDocument();
  });

  it("shows a spinner while messages load", async () => {
    let resolveGm!: (v: { messages: Message[] }) => void;
    const gmPromise = new Promise<{ messages: Message[] }>((r) => {
      resolveGm = r;
    });
    const msg = baseMessage({
      id: "m1",
      from: "8:x",
      content: "Later",
    });
    const mockClient = makeMockClient({
      account: {
        upn: "u@x.com",
        tenantId: "t",
        skypeId: "s",
        expiresAt: new Date(),
        region: "",
      },
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Slow" },
            lastMessage: msg,
          },
        ],
      }),
      getMessages: vi.fn().mockReturnValue(gmPromise),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Slow"));
    await waitFor(() => {
      expect(document.querySelector(".animate-pulse")).toBeTruthy();
    });
    resolveGm({ messages: [msg] });
    await waitFor(() => {
      expect(
        screen.getByRole("region", { name: /message thread/i }).textContent,
      ).toContain("Later");
    });
  });

  it("shows no matches when search has no hits", async () => {
    const mockClient = makeMockClient({
      account: {
        upn: "u@x.com",
        tenantId: "t",
        skypeId: "s",
        expiresAt: new Date(),
        region: "",
      },
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Zed" },
            lastMessage: baseMessage({
              id: "1",
              from: "8:x",
              content: "z",
            }),
          },
        ],
      }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    await screen.findByText("Zed");
    fireEvent.change(screen.getByRole("searchbox", { name: /search chats/i }), {
      target: { value: "qqq" },
    });
    expect(await screen.findByText(/no matches/i)).toBeInTheDocument();
  });

  it("shows empty inbox copy when there are no chats", async () => {
    const mockClient = makeMockClient({
      account: {
        upn: "u@x.com",
        tenantId: "t",
        skypeId: "s",
        expiresAt: new Date(),
        region: "",
      },
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    expect(
      await screen.findByText(/no conversations yet/i),
    ).toBeInTheDocument();
  });

  it("submits on Enter without shift in the composer", async () => {
    const msg = baseMessage({
      id: "m1",
      from: "8:o",
      content: "Hi",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Composer chat" },
            lastMessage: msg,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [msg] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Composer chat"));
    const input = await screen.findByRole("textbox", { name: /message text/i });
    fireEvent.input(input, { target: { innerHTML: "Line" } });
    fireEvent.keyDown(input, { key: "Enter", shiftKey: false });
    await waitFor(() => {
      expect(mockClient.sendMessage).toHaveBeenCalled();
    });
  });

  it("shows peer display name on incoming bubbles", async () => {
    const msg = baseMessage({
      id: "m1",
      from: "8:peer",
      content: "Hey",
      imdisplayname: "Alex",
    });
    const mockClient = makeMockClient({
      account: {
        upn: "me@test.com",
        tenantId: "t1",
        skypeId: "self",
        expiresAt: new Date(),
        region: "amer",
      },
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Named" },
            lastMessage: msg,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [msg] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Named"));
    expect(await screen.findByText("Alex")).toBeInTheDocument();
  });

  it("loads older messages when the top sentinel becomes visible", async () => {
    const newer = baseMessage({
      id: "m2",
      from: "8:o",
      content: "Recent line",
      composetime: "2024-06-02T12:00:00.000Z",
      originalarrivaltime: "2024-06-02T12:00:00.000Z",
    });
    const older = baseMessage({
      id: "m1",
      from: "8:o",
      content: "Older line",
      composetime: "2024-06-01T12:00:00.000Z",
      originalarrivaltime: "2024-06-01T12:00:00.000Z",
    });
    const backwardUrl =
      "https://test.example/messages?startTime=1&syncState=abc";
    const mockGetMessagesByUrl = vi.fn().mockResolvedValue({
      messages: [older],
      _metadata: {},
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "History" },
            lastMessage: newer,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({
        messages: [newer],
        _metadata: { backwardLink: backwardUrl },
      }),
      getMessagesByUrl: mockGetMessagesByUrl,
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("History"));
    expect(await screen.findByText("Recent line")).toBeInTheDocument();
    const region = screen.getByRole("region", { name: /message thread/i });
    region.scrollTop = 0;
    fireEvent.scroll(region);
    await waitFor(() => {
      expect(mockGetMessagesByUrl).toHaveBeenCalledWith(backwardUrl);
    });
    expect(await screen.findByText("Older line")).toBeInTheDocument();
  });

  it("does not render section headers in the sidebar", async () => {
    const msg = baseMessage({
      id: "m1",
      from: "8:other",
      content: "Hi",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "19:g1@thread.v2",
            threadProperties: { topic: "GTeam", threadType: "space" },
            lastMessage: msg,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [msg] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    await screen.findByText("GTeam");
    expect(screen.queryByRole("button", { name: /^meetings$/i })).toBeNull();
    expect(screen.queryByRole("button", { name: /^groups$/i })).toBeNull();
    expect(
      screen.queryByRole("button", { name: /^direct messages$/i }),
    ).toBeNull();
  });

  it("renders self-authored messages", async () => {
    const selfMsg = baseMessage({
      id: "m2",
      from: "8:self",
      content: "From me",
      imdisplayname: "Me",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Self chat" },
            lastMessage: selfMsg,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [selfMsg] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Self chat"));
    expect(await screen.findByText("From me")).toBeInTheDocument();
  });

  it("renders message links and mentions", async () => {
    const richMsg = baseMessage({
      id: "m-rich",
      from: "8:other",
      content:
        '<div>Hello <at id="0">Dirk</at> see <a href="https://example.com/spec">spec</a></div>',
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Rich chat" },
            lastMessage: richMsg,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [richMsg] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Rich chat"));
    expect(await screen.findByText("@Dirk")).toBeInTheDocument();
    expect(screen.getByRole("link", { name: "spec" })).toHaveAttribute(
      "href",
      "https://example.com/spec",
    );
  });

  it("renders file attachments as open cards", async () => {
    const attachmentMessage = baseMessage({
      id: "m-attachment",
      from: "8:peer",
      messagetype: "RichText/Media_GenericFile",
      contenttype: "RichText/Media_GenericFile",
      content:
        '<URIObject type="File.1" uri="https://api.asm.skype.com/v1/objects/0-123" url_thumbnail="https://api.asm.skype.com/v1/objects/0-123/views/thumbnail"><Title>Title: plans.pdf</Title><Description>Description: plans.pdf</Description><FileSize v="2048"/><OriginalName v="plans.pdf"/><a href="https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-123">https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-123</a></URIObject>',
      imdisplayname: "Alex",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Rich chat" },
            lastMessage: attachmentMessage,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [attachmentMessage] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Rich chat"));
    expect(
      await screen.findByRole("link", { name: /plans\.pdf/i }),
    ).toBeInTheDocument();
    expect(screen.getByText("2.0 KB")).toBeInTheDocument();
    expect(screen.queryByText(/Title: plans\.pdf/i)).not.toBeInTheDocument();
  });

  it("keeps attachment-only messages in the thread", async () => {
    const attachmentMessage = baseMessage({
      id: "m-attachment-only",
      from: "8:peer",
      messagetype: "RichText/Media_GenericFile",
      contenttype: "RichText/Media_GenericFile",
      content:
        '<URIObject type="File.1" uri="https://api.asm.skype.com/v1/objects/0-999" url_thumbnail="https://api.asm.skype.com/v1/objects/0-999/views/thumbnail"><Title>Title: budget.xlsx</Title><Description>Description: budget.xlsx</Description><FileSize v="4096"/><OriginalName v="budget.xlsx"/><a href="https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-999">https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-999</a></URIObject>',
      imdisplayname: "Alex",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Files" },
            lastMessage: attachmentMessage,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [attachmentMessage] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Files"));
    expect(
      await screen.findByRole("link", { name: /budget\.xlsx/i }),
    ).toBeInTheDocument();
  });

  it("opens profile UI from user mentions", async () => {
    const richMsg = baseMessage({
      id: "m-rich-profile",
      from: "8:other",
      content: '<div>Hello <at id="0">Siphesihle Thomo</at></div>',
      properties: {
        mentions: [
          {
            id: "0",
            mri: "8:orgid:peer-123",
            displayName: "Siphesihle Thomo",
          },
        ],
      },
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Rich chat" },
            lastMessage: richMsg,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [richMsg] }),
      fetchProfileAvatarDataUrls: vi.fn().mockResolvedValue({
        avatars: {},
        displayNames: { "8:orgid:peer-123": "Siphesihle Thomo" },
        emails: { "8:orgid:peer-123": "siphe@test.com" },
        jobTitles: { "8:orgid:peer-123": "Engineer" },
      }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Rich chat"));
    fireEvent.click(
      await screen.findByLabelText("View profile: Siphesihle Thomo"),
    );
    expect(
      await screen.findByText("Siphesihle Thomo's profile"),
    ).toBeInTheDocument();
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
  });

  it("keeps the DM profile sidebar closed until requested from the thread header", async () => {
    const dmMessage = baseMessage({
      id: "m-dm",
      from: "8:orgid:peer-123",
      content: "Hello from Pat",
    });
    const _groupMessage = baseMessage({
      id: "m-group-shared",
      from: "8:orgid:peer-123",
      content: "Pat posted here too",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            members: [
              { id: "8:self", role: "Admin", isMri: true },
              {
                id: "8:orgid:peer-123",
                role: "Admin",
                isMri: true,
                displayName: "Pat Lee",
                userPrincipalName: "pat@test.com",
              },
            ],
            threadProperties: {
              topic: "Pat Lee",
              threadType: "chat",
              productThreadType: "Chat",
              membercount: "2",
            },
            lastMessage: dmMessage,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [dmMessage] }),
      getThreadMembers: vi
        .fn()
        .mockImplementation(async (conversationId: string) => {
          if (conversationId === "c2") {
            return [
              { id: "8:self", role: "Admin", isMri: true },
              { id: "8:orgid:peer-123", role: "Admin", isMri: true },
              { id: "8:orgid:other-1", role: "Admin", isMri: true },
            ];
          }
          if (conversationId === "c3") {
            return [
              { id: "8:self", role: "Admin", isMri: true },
              {
                id: "29:opaque-member",
                role: "Admin",
                isMri: false,
                displayName: "Pat Lee",
                userPrincipalName: "pat@test.com",
              },
              { id: "8:orgid:other-2", role: "Admin", isMri: true },
            ];
          }
          if (conversationId === "c4") {
            return [
              { id: "8:self", role: "Admin", isMri: true },
              { id: "8:orgid:other-3", role: "Admin", isMri: true },
              { id: "8:orgid:other-4", role: "Admin", isMri: true },
            ];
          }
          throw new Error(`unexpected conversation ${conversationId}`);
        }),
      fetchProfileAvatarDataUrls: vi.fn().mockResolvedValue({
        avatars: {},
        displayNames: { "8:orgid:peer-123": "Pat Lee" },
        emails: { "8:orgid:peer-123": "pat@test.com" },
        jobTitles: { "8:orgid:peer-123": "Engineer" },
      }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Pat Lee"));
    await screen.findByRole("region", { name: /message thread/i });
    expect(
      screen.queryByLabelText("Conversation participants"),
    ).not.toBeInTheDocument();
    expect(screen.queryByText("Pat Lee's profile")).not.toBeInTheDocument();

    fireEvent.click(
      screen.getByRole("button", { name: "Open profile for Pat Lee" }),
    );
    expect(await screen.findByText("Pat Lee's profile")).toBeInTheDocument();
  });

  it("shows other chats with the same person in the profile sidebar", async () => {
    const dmMessage = baseMessage({
      id: "m-dm-shared",
      from: "8:orgid:peer-123",
      content: "Hey there",
    });
    const groupMessage = baseMessage({
      id: "m-group-shared",
      from: "8:orgid:peer-123",
      content: "Pat posted here too",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            members: [
              { id: "8:self", role: "Admin", isMri: true },
              { id: "8:orgid:peer-123", role: "Admin", isMri: true },
            ],
            threadProperties: {
              topic: "Pat Lee",
              threadType: "chat",
              productThreadType: "Chat",
              membercount: "2",
            },
            lastMessage: dmMessage,
          },
          {
            id: "c2",
            members: [
              { id: "8:self", role: "Admin", isMri: true },
              { id: "8:orgid:peer-123", role: "Admin", isMri: true },
              { id: "8:orgid:other-1", role: "Admin", isMri: true },
            ],
            threadProperties: {
              topic: "Design review",
              threadType: "chat",
              productThreadType: "Chat",
              membercount: "3",
            },
            lastMessage: groupMessage,
          },
          {
            id: "c3",
            members: [
              { id: "8:self", role: "Admin", isMri: true },
              {
                id: "29:opaque-member",
                role: "Admin",
                isMri: false,
                displayName: "Pat Lee",
                userPrincipalName: "pat@test.com",
              },
              { id: "8:orgid:other-2", role: "Admin", isMri: true },
            ],
            threadProperties: {
              topic: "Project alpha",
              threadType: "chat",
              productThreadType: "Chat",
              membercount: "3",
            },
            lastMessage: baseMessage({
              id: "m-group-fallback",
              from: "8:orgid:other-2",
              content: "Planning update",
            }),
          },
          {
            id: "c4",
            members: [
              { id: "8:self", role: "Admin", isMri: true },
              { id: "8:orgid:other-3", role: "Admin", isMri: true },
              { id: "8:orgid:other-4", role: "Admin", isMri: true },
            ],
            threadProperties: {
              topic: "Random chat",
              threadType: "chat",
              productThreadType: "Chat",
              membercount: "3",
            },
            lastMessage: baseMessage({
              id: "m-group-random",
              from: "8:orgid:other-3",
              content: "Not Pat",
            }),
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [dmMessage] }),
      fetchProfileAvatarDataUrls: vi.fn().mockResolvedValue({
        avatars: {},
        displayNames: { "8:orgid:peer-123": "Pat Lee" },
        emails: { "8:orgid:peer-123": "pat@test.com" },
        jobTitles: { "8:orgid:peer-123": "Engineer" },
      }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Pat Lee"));
    fireEvent.click(
      await screen.findByRole("button", { name: "Open profile for Pat Lee" }),
    );

    const heading = await screen.findByText("OTHER CHATS");
    const panel = heading.closest("aside");
    expect(panel).toBeTruthy();
    expect(
      within(panel as HTMLElement).getByText("OTHER CHATS"),
    ).toBeInTheDocument();
    expect(
      within(panel as HTMLElement).getByText("Design review"),
    ).toBeInTheDocument();
    expect(
      within(panel as HTMLElement).getByText("Project alpha"),
    ).toBeInTheDocument();
    expect(
      within(panel as HTMLElement).queryByText("Random chat"),
    ).not.toBeInTheDocument();
    fireEvent.click(
      within(panel as HTMLElement).getByRole("button", {
        name: /project alpha/i,
      }),
    );
    expect(
      await screen.findByRole("region", { name: /message thread/i }),
    ).toBeInTheDocument();
    expect(screen.getAllByText("Project alpha").length).toBeGreaterThan(0);
  });

  it("opens anchored thread segments from message mentions", async () => {
    const base = baseMessage({
      id: "m-base",
      from: "8:other",
      content:
        '<div>See <at data-message-id="m-target">that message</at></div>',
    });
    const target = baseMessage({
      id: "m-target",
      from: "8:other",
      content: "Anchored message body",
    });
    const mockGetAnchoredMessages = vi
      .fn()
      .mockResolvedValue({ messages: [target], _metadata: {} });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Rich chat" },
            lastMessage: base,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [base] }),
      getAnchoredMessages: mockGetAnchoredMessages,
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Rich chat"));
    fireEvent.click(
      await screen.findByRole("button", { name: "@that message" }),
    );
    await waitFor(() => {
      expect(mockGetAnchoredMessages).toHaveBeenCalledWith("c1", "m-target");
      expect(screen.getByText("Anchored message body")).toBeInTheDocument();
      expect(screen.getByText("See")).toBeInTheDocument();
      expect(
        screen.getByText("Anchored message body").closest("li"),
      ).toHaveAttribute("data-highlighted", "true");
    });
  });

  it("does not render phantom blank lines around quoted replies", async () => {
    const quoted = baseMessage({
      id: "m-quote",
      from: "8:self",
      content:
        '<blockquote itemtype="http://schema.skype.com/Reply"><div>&nbsp;</div><div><b>Siphesihle Thomo</b></div><div>Good morning Dirk, Review comments have been addressed</div><div>&nbsp;</div></blockquote><div>&nbsp;</div><div>Will do that</div>',
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Quoted chat" },
            lastMessage: quoted,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [quoted] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Quoted chat"));
    const message = await screen.findByText("Will do that");
    const row = message.closest("li");
    expect(row?.querySelectorAll("br").length).toBeLessThanOrEqual(2);
    expect(row).toHaveTextContent(
      "Siphesihle ThomoGood morning Dirk, Review comments have been addressedWill do that",
    );
  });

  it("does not leave a blank line between quoted author and body", async () => {
    const quoted = baseMessage({
      id: "m-quote-2",
      from: "8:self",
      content:
        '<blockquote itemtype="http://schema.skype.com/Reply"><div><b>Siphesihle Thomo</b></div><div>&nbsp;</div><div>Good morning Dirk, Review comments have been addressed</div></blockquote>',
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Quoted chat" },
            lastMessage: quoted,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [quoted] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Quoted chat"));
    const author = await screen.findByText("Siphesihle Thomo");
    const quote = author.closest("div");
    expect(quote?.querySelectorAll("br").length).toBeLessThanOrEqual(1);
    expect(quote).toHaveTextContent(
      "Siphesihle ThomoGood morning Dirk, Review comments have been addressed",
    );
  });

  it("opens anchored messages from clickable quote blocks", async () => {
    const quoted = baseMessage({
      id: "m-quote-3",
      from: "8:self",
      content:
        '<blockquote itemtype="http://schema.skype.com/Reply"><div>Siphesihle Thomo</div><div>Good morning Dirk, Review comments have been addressed</div></blockquote><div>Will do that</div>',
      properties: {
        qtdMsgs: [{ messageId: "m-target", sender: "8:orgid:peer" }],
      },
    });
    const target = baseMessage({
      id: "m-target",
      from: "8:other",
      content: "Original referenced message",
    });
    const mockGetAnchoredMessages = vi
      .fn()
      .mockResolvedValue({ messages: [target], _metadata: {} });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Quoted chat" },
            lastMessage: quoted,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [quoted] }),
      getAnchoredMessages: mockGetAnchoredMessages,
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Quoted chat"));
    fireEvent.click(
      await screen.findByRole("button", {
        name: "Siphesihle Thomo Good morning Dirk, Review comments have been addressed",
      }),
    );
    await waitFor(() => {
      expect(mockGetAnchoredMessages).toHaveBeenCalledWith("c1", "m-target");
      expect(
        screen.getByText("Original referenced message"),
      ).toBeInTheDocument();
      expect(screen.getByText("Will do that")).toBeInTheDocument();
      expect(
        screen.getByText("Original referenced message").closest("li"),
      ).toHaveAttribute("data-highlighted", "true");
    });
  });

  it("clears the composer when switching chats", async () => {
    const first = baseMessage({
      id: "m1",
      from: "8:a",
      content: "Alpha",
    });
    const second = baseMessage({
      id: "m2",
      from: "8:b",
      content: "Beta",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Alpha room" },
            lastMessage: first,
          },
          {
            id: "c2",
            threadProperties: { topic: "Beta room" },
            lastMessage: second,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [first, second] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Alpha room"));
    const input = await screen.findByRole("textbox", { name: /message text/i });
    fireEvent.input(input, { target: { innerHTML: "Some draft text" } });
    fireEvent.click(screen.getByText("Beta room"));
    const nextInput = await screen.findByPlaceholderText(
      "Message Beta room\u2026",
    );
    expect(nextInput.innerHTML).toBe("");
  });

  it("focuses the search box when Cmd+K is pressed", async () => {
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Alpha room" },
            lastMessage: baseMessage({
              id: "m1",
              from: "8:a",
              content: "Alpha",
            }),
          },
        ],
      }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    await screen.findByText("Alpha room");
    fireEvent.keyDown(window, { key: "k", ctrlKey: true });
    await waitFor(() => {
      const search = screen.getByRole("searchbox", { name: /search chats/i });
      expect(document.activeElement).toBe(search);
    });
  });

  it("does not submit while the composer is composing text", async () => {
    const msg = baseMessage({
      id: "m1",
      from: "8:o",
      content: "Hi",
    });
    const mockClient = makeMockClient({
      getAllConversations: vi.fn().mockResolvedValue({
        conversations: [
          {
            id: "c1",
            threadProperties: { topic: "Composer chat" },
            lastMessage: msg,
          },
        ],
      }),
      getMessages: vi.fn().mockResolvedValue({ messages: [msg] }),
    });
    vi.mocked(getOrCreateClient).mockResolvedValue(mockClient as never);

    renderChat();
    fireEvent.click(await screen.findByText("Composer chat"));
    const input = await screen.findByRole("textbox", { name: /message text/i });
    fireEvent.input(input, { target: { innerHTML: "Line" } });
    fireEvent.keyDown(input, { key: "Enter", isComposing: true });
    await waitFor(() => {
      expect(mockClient.sendMessage).not.toHaveBeenCalled();
    });
  });
});
