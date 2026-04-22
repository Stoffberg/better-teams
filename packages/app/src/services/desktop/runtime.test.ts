import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const invokeMock = vi.fn();

describe("electron-bridge", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-03-25T00:00:00Z"));
    vi.resetModules();
    vi.clearAllMocks();
    window.betterTeams = {
      teams: {
        extractTokens: vi.fn(),
        getAuthToken: vi.fn(),
        getAvailableAccounts: vi.fn(),
        getCachedPresence: invokeMock,
      },
      images: {
        cacheImageFile: vi.fn(),
        getCachedImageFile: vi.fn(),
        hasCachedImageFile: vi.fn(),
        filePathToAssetUrl: vi.fn((path: string) => `asset://${path}`),
      },
      http: {
        fetch: vi.fn(),
      },
      shell: {
        openExternal: vi.fn(),
      },
    };
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  it("reuses cached presence entries within the bridge ttl", async () => {
    invokeMock.mockResolvedValue([
      {
        mri: "8:orgid:one",
        presence: { availability: "Available", activity: "Active" },
      },
    ]);

    const { getCachedPresence } = await import("./runtime");

    const first = await getCachedPresence(["8:orgid:one"]);
    const second = await getCachedPresence(["8:orgid:one"]);

    expect(first).toEqual(second);
    expect(invokeMock).toHaveBeenCalledTimes(1);
  });

  it("shares one inflight presence request for identical mri batches", async () => {
    let resolveRequest: (
      value: Array<{ mri: string; presence: Record<string, string> }>,
    ) => void = () => {};
    const pendingRequest = new Promise<
      Array<{ mri: string; presence: Record<string, string> }>
    >((resolve) => {
      resolveRequest = resolve;
    });
    invokeMock.mockImplementation(() => pendingRequest);

    const { getCachedPresence } = await import("./runtime");

    const first = getCachedPresence(["8:orgid:one", "8:orgid:two"]);
    const second = getCachedPresence(["8:orgid:two", "8:orgid:one"]);

    resolveRequest([
      {
        mri: "8:orgid:one",
        presence: { availability: "Busy", activity: "InACall" },
      },
      {
        mri: "8:orgid:two",
        presence: { availability: "Available", activity: "Active" },
      },
    ]);

    const [firstResult, secondResult] = await Promise.all([first, second]);

    expect(firstResult).toEqual(secondResult);
    expect(invokeMock).toHaveBeenCalledTimes(1);
  });

  it("refreshes presence after the bridge ttl expires", async () => {
    invokeMock
      .mockResolvedValueOnce([
        {
          mri: "8:orgid:one",
          presence: { availability: "Available", activity: "Active" },
        },
      ])
      .mockResolvedValueOnce([
        {
          mri: "8:orgid:one",
          presence: { availability: "Away", activity: "Away" },
        },
      ]);

    const { getCachedPresence } = await import("./runtime");

    await getCachedPresence(["8:orgid:one"]);
    vi.advanceTimersByTime(15_001);
    const refreshed = await getCachedPresence(["8:orgid:one"]);

    expect(refreshed["8:orgid:one"]).toEqual({
      availability: "Away",
      activity: "Away",
    });
    expect(invokeMock).toHaveBeenCalledTimes(2);
  });
});
