import { beforeEach, describe, expect, it, vi } from "vitest";
import { SqliteThreadCache } from "@/lib/sqlite-cache";
import { getOrCreateClient } from "@/lib/teams-client-factory";
import { preloadConversationThread } from "./teams-thread-preload";

vi.mock("@/lib/sqlite-cache", () => ({
  SqliteThreadCache: {
    getFreshSnapshot: vi.fn(),
    storeThread: vi.fn().mockResolvedValue(undefined),
  },
}));

vi.mock("@/lib/teams-client-factory", () => ({
  getOrCreateClient: vi.fn(),
}));

describe("preloadConversationThread", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("returns a fresh cached thread without refetching", async () => {
    vi.mocked(SqliteThreadCache.getFreshSnapshot).mockResolvedValue({
      data: {
        messages: [{ id: "m1" }],
        olderPageUrl: null,
        moreOlder: false,
      },
      updatedAt: Date.now(),
    } as never);

    const result = await preloadConversationThread("t1", "c1");

    expect(result.messages).toEqual([{ id: "m1" }]);
    expect(getOrCreateClient).not.toHaveBeenCalled();
    expect(SqliteThreadCache.storeThread).not.toHaveBeenCalled();
  });

  it("fetches and stores the thread when the cache is cold", async () => {
    vi.mocked(SqliteThreadCache.getFreshSnapshot).mockResolvedValue(null);
    vi.mocked(getOrCreateClient).mockResolvedValue({
      getMessages: vi.fn().mockResolvedValue({
        messages: [
          {
            id: "m2",
            composetime: "2024-06-01T12:00:00.000Z",
            originalarrivaltime: "2024-06-01T12:00:00.000Z",
          },
        ],
      }),
    } as never);

    const result = await preloadConversationThread("t1", "c1");

    expect(getOrCreateClient).toHaveBeenCalledWith("t1");
    expect(SqliteThreadCache.storeThread).toHaveBeenCalledWith("t1", "c1", {
      messages: [
        {
          id: "m2",
          composetime: "2024-06-01T12:00:00.000Z",
          originalarrivaltime: "2024-06-01T12:00:00.000Z",
        },
      ],
      olderPageUrl: null,
      moreOlder: false,
    });
    expect(result.messages[0]?.id).toBe("m2");
  });
});
