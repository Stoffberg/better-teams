import { beforeEach, describe, expect, it, vi } from "vitest";
import { getOrCreateClient } from "@/lib/teams-client-factory";
import { preloadConversationThread } from "./teams-thread-preload";

vi.mock("@/lib/teams-client-factory", () => ({
  getOrCreateClient: vi.fn(),
}));

describe("preloadConversationThread", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("fetches the first page from Teams", async () => {
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
    expect(result.messages[0]?.id).toBe("m2");
  });
});
