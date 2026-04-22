import { getCachedMessages } from "@better-teams/app/services/desktop/runtime";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { getTeamsClient } from "./client";
import { teamsThreadService } from "./thread";

vi.mock("@better-teams/app/services/desktop/runtime", () => ({
  getCachedMessages: vi.fn().mockResolvedValue(null),
}));

vi.mock("./client", () => ({
  getTeamsClient: vi.fn(),
}));

const rawMessage = (
  id: string,
  from: string,
  imdisplayname: string,
  content: string,
) => ({
  id,
  conversationId: "19:cached-members@thread.v2",
  type: "Message",
  messagetype: "Text",
  contenttype: "text",
  from,
  imdisplayname,
  content,
  composetime: "2024-06-01T12:00:00.000Z",
  originalarrivaltime: "2024-06-01T12:00:00.000Z",
});

describe("teamsThreadService", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(getCachedMessages).mockResolvedValue(null);
  });

  it("uses cached message authors for members when the roster endpoint fails", async () => {
    vi.mocked(getTeamsClient).mockResolvedValue({
      getThreadMembers: vi.fn().mockRejectedValue(new Error("offline")),
    } as never);
    vi.mocked(getCachedMessages).mockResolvedValue({
      messages: [
        rawMessage("m-self", "8:self", "Dirk Beukes", "Checking this"),
        rawMessage("m-peer-a", "8:orgid:peer-a", "Martha Dorey", "Looks fine"),
        rawMessage("m-peer-b", "8:orgid:peer-b", "Matt Amphlett", "Agreed"),
        rawMessage("m-agent", "28:agent-service", "Agent Service", "Ignore me"),
        rawMessage("m-duplicate", "8:orgid:peer-b", "Matt Amphlett", "Again"),
      ],
    });

    await expect(
      teamsThreadService.getMembers("tenant-1", "19:cached-members@thread.v2"),
    ).resolves.toEqual([
      {
        id: "8:self",
        role: "User",
        isMri: true,
        displayName: "Dirk Beukes",
      },
      {
        id: "8:orgid:peer-a",
        role: "User",
        isMri: true,
        displayName: "Martha Dorey",
      },
      {
        id: "8:orgid:peer-b",
        role: "User",
        isMri: true,
        displayName: "Matt Amphlett",
      },
    ]);
  });

  it("filters agent members from live rosters", async () => {
    vi.mocked(getTeamsClient).mockResolvedValue({
      getThreadMembers: vi.fn().mockResolvedValue([
        { id: "8:self", role: "Admin", isMri: true, displayName: "Dirk" },
        {
          id: "8:orgid:peer",
          role: "Admin",
          isMri: true,
          displayName: "Martha",
        },
        {
          id: "28:agent-service",
          role: "Admin",
          isMri: false,
          displayName: "Agent Service",
        },
      ]),
    } as never);

    await expect(
      teamsThreadService.getMembers("tenant-1", "19:live-members@thread.v2"),
    ).resolves.toEqual([
      { id: "8:self", role: "Admin", isMri: true, displayName: "Dirk" },
      {
        id: "8:orgid:peer",
        role: "Admin",
        isMri: true,
        displayName: "Martha",
      },
    ]);
  });
});
