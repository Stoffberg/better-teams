import { describe, expect, it } from "vitest";
import { teamsKeys } from "@/lib/teams-query-keys";
import { shouldPersistQuery } from "./QueryProvider";

describe("shouldPersistQuery", () => {
  it("does not persist teams cache entries that are already backed by sqlite", () => {
    expect(
      shouldPersistQuery({
        queryKey: teamsKeys.profileAvatars("tenant-1", "mri-a"),
      }),
    ).toBe(false);
    expect(
      shouldPersistQuery({
        queryKey: teamsKeys.thread("tenant-1", "conversation-1"),
      }),
    ).toBe(false);
  });

  it("still persists non teams queries", () => {
    expect(
      shouldPersistQuery({
        queryKey: ["settings-dialog"],
      }),
    ).toBe(true);
  });

  it("does not persist transient conversation open requests", () => {
    expect(
      shouldPersistQuery({
        queryKey: ["open-conversation-request"],
      }),
    ).toBe(false);
  });
});
