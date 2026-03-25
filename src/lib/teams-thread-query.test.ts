import { describe, expect, it } from "vitest";
import type { Message } from "@/services/teams/types";
import { threadQueryDataFromResponse } from "./teams-thread-query";

function m(partial: Partial<Message> & Pick<Message, "id" | "from">): Message {
  return {
    conversationId: "c",
    type: "Message",
    messagetype: "Text",
    contenttype: "text",
    content: "x",
    composetime: "2024-01-01T00:00:00.000Z",
    originalarrivaltime: "2024-01-01T00:00:00.000Z",
    ...partial,
  };
}

describe("threadQueryDataFromResponse", () => {
  it("reverses to ascending order and exposes backwardLink for pagination", () => {
    const newer = m({
      id: "b",
      from: "8:x",
      originalarrivaltime: "2024-01-02T00:00:00.000Z",
    });
    const older = m({
      id: "a",
      from: "8:x",
      originalarrivaltime: "2024-01-01T00:00:00.000Z",
    });
    const out = threadQueryDataFromResponse({
      messages: [newer, older],
      _metadata: { backwardLink: "https://example.com/messages?syncState=abc" },
    });
    expect(out.messages.map((x) => x.id)).toEqual(["a", "b"]);
    expect(out.olderPageUrl).toBe("https://example.com/messages?syncState=abc");
    expect(out.moreOlder).toBe(true);
  });

  it("sets moreOlder false when no backwardLink", () => {
    const msg = m({ id: "a", from: "8:x" });
    const out = threadQueryDataFromResponse({
      messages: [msg],
      _metadata: {},
    });
    expect(out.olderPageUrl).toBeNull();
    expect(out.moreOlder).toBe(false);
  });
});
