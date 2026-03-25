import { describe, expect, it } from "vitest";
import { presenceBadgeClassName, presenceDescription } from "./teams-presence";

describe("teams-presence", () => {
  it("maps common presence states to badge colors", () => {
    expect(presenceBadgeClassName()).toBeNull();
    expect(
      presenceBadgeClassName({ availability: "Available", activity: "Active" }),
    ).toBe("bg-emerald-500");
    expect(
      presenceBadgeClassName({ availability: "Busy", activity: "InACall" }),
    ).toBe("bg-rose-500");
    expect(
      presenceBadgeClassName({
        availability: "DoNotDisturb",
        activity: "Presenting",
      }),
    ).toBe("bg-rose-600");
    expect(
      presenceBadgeClassName({ availability: "Away", activity: "Away" }),
    ).toBe("bg-amber-400");
    expect(
      presenceBadgeClassName({ availability: "Offline", activity: "Offline" }),
    ).toBe("bg-zinc-400");
  });

  it("includes status messages in the description", () => {
    expect(
      presenceDescription({
        availability: "Available",
        activity: "Active",
        statusMessage: {
          message: "Heads down",
          expiry: "2026-03-25T10:00:00Z",
        },
      }),
    ).toBe("Available: Heads down");
  });
});
