import { describe, expect, it } from "vitest";
import type { Message } from "@/services/teams/types";
import {
  captureScrollRestoreAnchor,
  mergeThreadSnapshots,
  olderPrefetchThresholdForVelocity,
  profileMessageConversationId,
  restoreScrollRestoreAnchor,
  shouldPrefetchOlderMessages,
} from "./ThreadView";

describe("profileMessageConversationId", () => {
  it("hides the message action when the profile is opened inside a DM", () => {
    expect(
      profileMessageConversationId("dm", "c1", [
        { id: "c2", title: "Design review", kind: "group" },
      ]),
    ).toBeUndefined();
  });

  it("prefers an existing DM when the profile is opened from a group or meeting", () => {
    expect(
      profileMessageConversationId("group", "c2", [
        { id: "c3", title: "Design review", kind: "group" },
        { id: "c1", title: "Pat Lee", kind: "dm" },
      ]),
    ).toBe("c1");
  });

  it("returns undefined when there is no DM to open", () => {
    expect(
      profileMessageConversationId("meeting", "c2", [
        { id: "c3", title: "Design review", kind: "group" },
      ]),
    ).toBeUndefined();
  });
});

describe("scroll restore anchors", () => {
  it("captures the first visible message instead of a raw scroll position", () => {
    const viewport = document.createElement("div");
    const hidden = document.createElement("div");
    const anchor = document.createElement("div");

    hidden.dataset.messageId = "hidden";
    anchor.dataset.messageId = "anchor";
    viewport.append(hidden, anchor);

    Object.defineProperty(viewport, "scrollHeight", { value: 1200 });
    Object.defineProperty(viewport, "scrollTop", {
      value: 96,
      writable: true,
    });
    viewport.getBoundingClientRect = () => ({ top: 200 }) as DOMRect;
    hidden.getBoundingClientRect = () => ({ top: 140, bottom: 180 }) as DOMRect;
    anchor.getBoundingClientRect = () => ({ top: 228, bottom: 280 }) as DOMRect;

    expect(captureScrollRestoreAnchor(viewport)).toEqual({
      messageId: "anchor",
      scrollHeight: 1200,
      scrollTop: 96,
      top: 28,
    });
  });

  it("restores against the same anchor after prepending older messages", () => {
    const viewport = document.createElement("div");
    const anchor = document.createElement("div");

    anchor.dataset.messageId = "anchor";
    viewport.append(anchor);

    Object.defineProperty(viewport, "scrollHeight", {
      value: 1600,
      writable: true,
    });
    Object.defineProperty(viewport, "scrollTop", {
      value: 0,
      writable: true,
    });
    viewport.getBoundingClientRect = () => ({ top: 120 }) as DOMRect;
    anchor.getBoundingClientRect = () => ({ top: 182, bottom: 240 }) as DOMRect;

    restoreScrollRestoreAnchor(viewport, {
      messageId: "anchor",
      scrollHeight: 1200,
      scrollTop: 96,
      top: 28,
    });

    expect(viewport.scrollTop).toBe(34);
  });
});

describe("older message prefetch", () => {
  it("starts fetching before the viewport reaches the top edge", () => {
    expect(shouldPrefetchOlderMessages(1799)).toBe(true);
    expect(shouldPrefetchOlderMessages(1800)).toBe(true);
    expect(shouldPrefetchOlderMessages(1801)).toBe(false);
  });

  it("pulls earlier when the user is scrolling upward quickly", () => {
    expect(olderPrefetchThresholdForVelocity(0)).toBe(1800);
    expect(olderPrefetchThresholdForVelocity(1)).toBe(3300);
    expect(olderPrefetchThresholdForVelocity(10)).toBe(7800);
    expect(shouldPrefetchOlderMessages(3200, 2)).toBe(true);
    expect(shouldPrefetchOlderMessages(3200, 0.2)).toBe(false);
  });
});

describe("thread snapshot merge", () => {
  it("keeps cached older messages when live data arrives", () => {
    const oldMessage = {
      id: "m-1",
      from: "8:a",
      conversationId: "c1",
      originalarrivaltime: "2026-04-21T10:00:00.000Z",
    } as Message;
    const liveMessage = {
      id: "m-2",
      from: "8:a",
      conversationId: "c1",
      originalarrivaltime: "2026-04-21T10:05:00.000Z",
    } as Message;

    expect(
      mergeThreadSnapshots(
        {
          messages: [liveMessage],
          olderPageUrl: "live-older",
          moreOlder: true,
        },
        {
          messages: [oldMessage, liveMessage],
          olderPageUrl: "cached-older",
          moreOlder: true,
        },
      ),
    ).toEqual({
      messages: [oldMessage, liveMessage],
      olderPageUrl: "live-older",
      moreOlder: true,
    });
  });
});
