import { describe, expect, it } from "vitest";
import { conversationChatKind, messagePlainText } from "./chat";
import type { Conversation } from "./teams/types";

describe("chat compatibility exports", () => {
  it("keeps legacy chat helpers available through the package chat entrypoint", () => {
    const conversation: Conversation = {
      id: "19:thread@thread.v2",
      threadProperties: { membercount: "2" },
    };

    expect(conversationChatKind(conversation)).toBe("dm");
    expect(messagePlainText("<b>Hello</b><br>world")).toContain("Hello");
  });
});
