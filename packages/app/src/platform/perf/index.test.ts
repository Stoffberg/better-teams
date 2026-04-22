import { describe, expect, it } from "vitest";
import {
  beginPerfMeasure,
  PERF_FLAG,
  recordPerfMetric,
  resetPerfStore,
  updatePerfSnapshot,
} from "./index";

describe("perf store", () => {
  it("stays inactive when the perf flag is disabled", () => {
    localStorage.removeItem(PERF_FLAG);
    resetPerfStore();

    recordPerfMetric("workspace.load", { status: "ok" });
    updatePerfSnapshot("workspace", { conversations: 2 });
    const end = beginPerfMeasure("workspace.select");
    end({ status: "ok" });

    expect(window.__BETTER_TEAMS_PERF__).toBeUndefined();
  });

  it("collects metrics and snapshots when the perf flag is enabled", () => {
    localStorage.setItem(PERF_FLAG, "1");
    resetPerfStore();

    recordPerfMetric("workspace.load", { status: "ok" });
    updatePerfSnapshot("workspace", { conversations: 2 });
    const end = beginPerfMeasure("workspace.select", { source: "sidebar" });
    end({ status: "ok" });

    expect(window.__BETTER_TEAMS_PERF__?.metrics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          name: "workspace.load",
          detail: { status: "ok" },
        }),
        expect.objectContaining({
          name: "workspace.select.start",
          detail: { source: "sidebar" },
        }),
        expect.objectContaining({
          name: "workspace.select",
          detail: { source: "sidebar", status: "ok" },
        }),
      ]),
    );
    expect(window.__BETTER_TEAMS_PERF__?.snapshots.workspace?.values).toEqual({
      conversations: 2,
    });
  });
});
