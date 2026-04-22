import { beforeEach, describe, expect, it, vi } from "vitest";

const fetchMock = vi.fn();
const initializeMock = vi.fn();
const apiClientConstructor = vi.fn();

vi.mock("@/lib/electron-fetch", () => ({
  fetch: fetchMock,
}));

vi.mock("@/services/teams/api-client", () => ({
  TeamsApiClient: vi.fn().mockImplementation(function TeamsApiClientMock() {
    apiClientConstructor();
    return {
      initialize: initializeMock,
    };
  }),
}));

describe("teams-client-factory", () => {
  beforeEach(() => {
    vi.resetModules();
    vi.clearAllMocks();
    initializeMock.mockResolvedValue(undefined);
  });

  it("shares one initialization per tenant across concurrent callers", async () => {
    const { getOrCreateClient } = await import("./teams-client-factory");

    const [first, second] = await Promise.all([
      getOrCreateClient("tenant-1"),
      getOrCreateClient("tenant-1"),
    ]);

    expect(first).toBe(second);
    expect(apiClientConstructor).toHaveBeenCalledTimes(1);
    expect(initializeMock).toHaveBeenCalledTimes(1);
  });

  it("clears an inflight initialization after failure so the next call can retry", async () => {
    initializeMock.mockRejectedValueOnce(new Error("boom"));

    const { clearClientCache, getOrCreateClient } = await import(
      "./teams-client-factory"
    );

    await expect(getOrCreateClient("tenant-1")).rejects.toThrow("boom");
    clearClientCache();
    await getOrCreateClient("tenant-1");

    expect(apiClientConstructor).toHaveBeenCalledTimes(2);
    expect(initializeMock).toHaveBeenCalledTimes(2);
  });
});
