import { describe, expect, it } from "vitest";
import type { Conversation, Message } from "@/services/teams/types";
import {
  conversationChatKind,
  conversationPreview,
  conversationTitle,
  filterConversationsForPipeline,
  formatDayLabel,
  formatDetailedTimestamp,
  formatMessageTime,
  formatThreadDayDividerLabel,
  gapBetweenMessages,
  includeConversationInSidebar,
  initialsFromLabel,
  isCallLogsStubConversation,
  isCompanyCommunicationsSidebarThread,
  isEditedMessage,
  isLikelySystemOrCallBlob,
  isRenderableChatMessage,
  isSelfMessage,
  messageBodyForDisplay,
  messagePartsForDisplay,
  messagePlainText,
  messageReadStatus,
  messageReadTimestamp,
  messageRichPartsForDisplay,
  messageTimestamp,
  normalizePreviewText,
  parseConsumptionHorizon,
  parseMessageQuoteAndBody,
  partitionConversationsByKind,
  sortConversationsByActivity,
  textLooksLikeTeamsCallLogStub,
} from "./chat-format";

function msg(
  partial: Partial<Message> & Pick<Message, "id" | "from">,
): Message {
  return {
    conversationId: "c",
    type: "Message",
    messagetype: "Text",
    contenttype: "text",
    content: "",
    composetime: "",
    originalarrivaltime: "",
    ...partial,
  };
}

describe("formatThreadDayDividerLabel", () => {
  it("uses long weekday and month for past dates", () => {
    const s = formatThreadDayDividerLabel("2024-03-20T12:00:00.000Z");
    expect(s.toLowerCase()).toContain("march");
    expect(s.toLowerCase()).toContain("20");
  });
});

describe("messageRichPartsForDisplay", () => {
  it("extracts file attachment metadata from URIObject messages", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "file-1",
        from: "8:peer",
        messagetype: "RichText/Media_GenericFile",
        contenttype: "RichText/Media_GenericFile",
        content:
          '<URIObject type="File.1" uri="https://api.asm.skype.com/v1/objects/0-123" url_thumbnail="https://api.asm.skype.com/v1/objects/0-123/views/thumbnail"><Title>Title: plans.pdf</Title><Description>Description: plans.pdf</Description><FileSize v="2048"/><OriginalName v="plans.pdf"/><a href="https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-123">https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-123</a></URIObject>',
      }),
    );

    expect(parts?.attachments).toEqual([
      {
        kind: "file",
        objectUrl: "https://api.asm.skype.com/v1/objects/0-123",
        openUrl:
          "https://login.skype.com/login/sso?go=webclient.xmm&docid=0-123",
        thumbnailUrl:
          "https://api.asm.skype.com/v1/objects/0-123/views/thumbnail",
        title: "plans.pdf",
        fileName: "plans.pdf",
        fileSize: 2048,
        fileExtension: "pdf",
      },
    ]);
    expect(parts?.body).toEqual([]);
  });

  it("extracts attachment from URIObject even when messagetype is plain Text", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "file-text",
        from: "8:peer",
        messagetype: "Text",
        contenttype: "text",
        content:
          '<URIObject type="File.1" uri="https://api.asm.skype.com/v1/objects/0-123"><Title>Title: three-body.html</Title><Description>Description: three-body.html</Description><OriginalName v="three-body.html"/><a href="https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-123">https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-123</a></URIObject>',
      }),
    );

    expect(parts?.attachments).toEqual([
      {
        kind: "file",
        objectUrl: "https://api.asm.skype.com/v1/objects/0-123",
        openUrl:
          "https://login.skype.com/login/sso?go=webclient.xmm&docid=0-123",
        title: "three-body.html",
        fileName: "three-body.html",
        fileExtension: "html",
      },
    ]);
    expect(parts?.body).toEqual([]);
  });
});

describe("conversationChatKind", () => {
  it("treats team channel thread as group when roster exceeds two", () => {
    const c: Conversation = {
      id: "19:abc@thread.tacv2",
      threadProperties: { membercount: "6" },
    };
    expect(conversationChatKind(c)).toBe("group");
  });

  it("treats tacv2 as dm when roster is two", () => {
    const c: Conversation = {
      id: "19:abc@thread.tacv2",
      threadProperties: { membercount: "2" },
    };
    expect(conversationChatKind(c)).toBe("dm");
  });

  it("treats more than two roster entries as group", () => {
    const c: Conversation = {
      id: "19:x@thread.v2",
      threadProperties: { membercount: "3" },
    };
    expect(conversationChatKind(c)).toBe("group");
  });

  it("treats Teams group chat as group when roster is more than two", () => {
    const c: Conversation = {
      id: "19:4c7d0247747f4d9da394d99eb9815e65@thread.v2",
      threadProperties: {
        topic: "Internal Engineering",
        threadType: "chat",
        productThreadType: "Chat",
        membercount: "4",
      },
    };
    expect(conversationChatKind(c)).toBe("group");
  });

  it("treats chat as group when members array length is more than two", () => {
    const c: Conversation = {
      id: "19:2c72a99f50e94b6e88fd1c19a622e93c@thread.skype",
      threadProperties: {
        topic: "Engineering",
        threadType: "chat",
        productThreadType: "Chat",
      },
      members: [
        { id: "8:a", role: "User", isMri: true },
        { id: "8:b", role: "User", isMri: true },
        { id: "8:c", role: "User", isMri: true },
      ],
    };
    expect(conversationChatKind(c)).toBe("group");
  });

  it("keeps short titled chat thread as dm", () => {
    const c: Conversation = {
      id: "19:x@thread.v2",
      threadProperties: {
        membercount: "2",
        topic: "Pat Lee",
        threadType: "chat",
        productThreadType: "Chat",
      },
    };
    expect(conversationChatKind(c)).toBe("dm");
  });

  it("uses max of membercount and members length for roster", () => {
    const c: Conversation = {
      id: "19:x@thread.v2",
      threadProperties: { membercount: "2", topic: "Squad" },
      members: [
        { id: "a", role: "User", isMri: true },
        { id: "b", role: "User", isMri: true },
        { id: "c", role: "User", isMri: true },
      ],
    };
    expect(conversationChatKind(c)).toBe("group");
  });

  it("uses threadProperties.memberCount when membercount understates", () => {
    const c: Conversation = {
      id: "19:x@thread.v2",
      threadProperties: {
        membercount: "2",
        memberCount: 12,
        topic: "Internal Engineering",
      },
    };
    expect(conversationChatKind(c)).toBe("group");
  });

  it("treats conversationType channel as dm when roster is two", () => {
    const c: Conversation = {
      id: "19:x@thread.v2",
      conversationType: "channel",
      threadProperties: { membercount: "2", topic: "Internal Engineering" },
    };
    expect(conversationChatKind(c)).toBe("dm");
  });

  it("treats channel-shaped thread as group when roster exceeds two", () => {
    const c: Conversation = {
      id: "19:821d63b920fa4deea228742d2107f26c@thread.tacv2",
      threadProperties: {
        topic: "Migrate Enterprise Apps to Microsoft Azure",
        threadType: "topic",
        productThreadType: "TeamsStandardChannel",
        membercount: "8",
      },
    };
    expect(conversationChatKind(c)).toBe("group");
  });

  it("treats consumer pair thread id as dm regardless of roster hints", () => {
    const c: Conversation = {
      id: "19:48bb84f7-e080-4d88-8043-ae0e8cf745fc_f4cc62d6-05d5-48b0-9feb-ffe47197d860@unq.gbl.spaces",
      threadProperties: { membercount: "99" },
    };
    expect(conversationChatKind(c)).toBe("dm");
  });

  it("treats sparse list row as group when not a consumer pair and roster is not exactly two", () => {
    const c: Conversation = {
      id: "19:4c7d0247747f4d9da394d99eb9815e65@thread.v2",
      threadProperties: {
        topic: "Internal Engineering",
        threadType: "chat",
        productThreadType: "Chat",
      },
    };
    expect(conversationChatKind(c)).toBe("group");
  });

  it("treats non-thread ids as dm when roster is unknown", () => {
    expect(conversationChatKind({ id: "c1" })).toBe("dm");
  });
});

describe("partitionConversationsByKind", () => {
  it("places Internal Engineering style list row into groups", () => {
    const ie: Conversation = {
      id: "19:4c7d0247747f4d9da394d99eb9815e65@thread.v2",
      threadProperties: {
        topic: "Internal Engineering",
        threadType: "chat",
        productThreadType: "Chat",
      },
    };
    const { groups, dms, meetings } = partitionConversationsByKind([ie]);
    expect(groups.map((c) => c.id)).toContain(ie.id);
    expect(dms).toHaveLength(0);
    expect(meetings).toHaveLength(0);
  });

  it("splits by inferred kind", () => {
    const dm: Conversation = {
      id: "d",
      threadProperties: { membercount: "2" },
    };
    const grp: Conversation = {
      id: "g",
      threadProperties: { membercount: "5", topic: "Squad" },
    };
    const meet: Conversation = {
      id: "m",
      threadProperties: { topic: "Meeting notes" },
    };
    const { meetings, groups, dms } = partitionConversationsByKind([
      dm,
      grp,
      meet,
    ]);
    expect(dms.map((c) => c.id)).toContain("d");
    expect(groups.map((c) => c.id)).toContain("g");
    expect(meetings.map((c) => c.id)).toContain("m");
  });
});

describe("initialsFromLabel", () => {
  it("uses two letters for a single word", () => {
    expect(initialsFromLabel("alex")).toBe("AL");
  });

  it("uses first letters of first and last word", () => {
    expect(initialsFromLabel("Alex Smith")).toBe("AS");
  });

  it("returns question mark for empty", () => {
    expect(initialsFromLabel("  ")).toBe("?");
  });
});

describe("conversationTitle", () => {
  it("prefers thread topic", () => {
    const c: Conversation = {
      id: "x",
      threadProperties: { topic: " Alpha " },
    };
    expect(conversationTitle(c)).toBe("Alpha");
  });

  it("strips trailing Play from topic", () => {
    const c: Conversation = {
      id: "x",
      threadProperties: { topic: "Sprint Review Play" },
    };
    expect(conversationTitle(c)).toBe("Sprint Review");
  });

  it("falls back to last sender display name", () => {
    const c: Conversation = {
      id: "x",
      lastMessage: msg({
        id: "1",
        from: "8:a",
        imdisplayname: "Pat",
      }),
    };
    expect(conversationTitle(c)).toBe("Pat");
  });

  it("dm uses peer member display name when last sender is self", () => {
    const c: Conversation = {
      id: "19:pair@unq.gbl.spaces",
      threadProperties: { membercount: "2" },
      members: [
        {
          id: "8:orgid:self-uuid",
          role: "User",
          isMri: true,
          displayName: "Dirk",
        },
        {
          id: "8:orgid:peer-uuid",
          role: "User",
          isMri: true,
          displayName: "Martha",
        },
      ],
      lastMessage: msg({
        id: "1",
        from: "8:orgid:self-uuid",
        imdisplayname: "Dirk",
      }),
    };
    expect(conversationTitle(c, "orgid:self-uuid")).toBe("Martha");
  });

  it("dm falls back to Direct message when self sent last and members lack names", () => {
    const c: Conversation = {
      id: "19:pair@unq.gbl.spaces",
      threadProperties: { membercount: "2" },
      members: [
        { id: "8:orgid:self-uuid", role: "User", isMri: true },
        { id: "8:orgid:peer-uuid", role: "User", isMri: true },
      ],
      lastMessage: msg({
        id: "1",
        from: "8:orgid:self-uuid",
        imdisplayname: "Dirk",
      }),
    };
    expect(conversationTitle(c, "orgid:self-uuid")).toBe("Direct message");
  });

  it("dm uses short-profile display name when self sent last and members lack names", () => {
    const c: Conversation = {
      id: "19:pair@unq.gbl.spaces",
      threadProperties: { membercount: "2" },
      members: [
        { id: "8:orgid:self-uuid", role: "User", isMri: true },
        { id: "8:orgid:peer-uuid", role: "User", isMri: true },
      ],
      lastMessage: msg({
        id: "1",
        from: "8:orgid:self-uuid",
        imdisplayname: "Dirk",
      }),
    };
    expect(conversationTitle(c, "orgid:self-uuid", "Martha")).toBe("Martha");
  });

  it("group without topic does not use last sender display name", () => {
    const c: Conversation = {
      id: "19:abc@thread.tacv2",
      threadProperties: { membercount: "5" },
      lastMessage: msg({
        id: "1",
        from: "8:other",
        imdisplayname: "Pat",
      }),
    };
    expect(conversationTitle(c)).toBe("Group chat");
  });

  it("group without topic uses member names with overflow count", () => {
    const c: Conversation = {
      id: "19:abc@thread.v2",
      threadProperties: { membercount: "6" },
      members: [
        {
          id: "8:orgid:self-uuid",
          displayName: "Dirk",
          role: "User",
          isMri: true,
        },
        {
          id: "8:orgid:abdul-uuid",
          displayName: "Abdul",
          role: "User",
          isMri: true,
        },
        {
          id: "8:orgid:daniel-uuid",
          displayName: "Daniel",
          role: "User",
          isMri: true,
        },
        {
          id: "8:orgid:kane-uuid",
          displayName: "Kane",
          role: "User",
          isMri: true,
        },
        {
          id: "8:orgid:pat-uuid",
          displayName: "Pat",
          role: "User",
          isMri: true,
        },
        {
          id: "8:orgid:zoe-uuid",
          displayName: "Zoe",
          role: "User",
          isMri: true,
        },
      ],
    };
    expect(conversationTitle(c, "orgid:self-uuid")).toBe(
      "Abdul, Daniel, Kane, +2",
    );
  });

  it("group without topic uses all member names when three or fewer peers exist", () => {
    const c: Conversation = {
      id: "19:abc@thread.v2",
      threadProperties: { membercount: "4" },
      members: [
        {
          id: "8:orgid:self-uuid",
          displayName: "Dirk",
          role: "User",
          isMri: true,
        },
        {
          id: "8:orgid:abdul-uuid",
          displayName: "Abdul",
          role: "User",
          isMri: true,
        },
        {
          id: "8:orgid:daniel-uuid",
          displayName: "Daniel",
          role: "User",
          isMri: true,
        },
        {
          id: "8:orgid:kane-uuid",
          displayName: "Kane",
          role: "User",
          isMri: true,
        },
      ],
    };
    expect(conversationTitle(c, "orgid:self-uuid")).toBe("Abdul, Daniel, Kane");
  });

  it("meeting without topic does not use last sender display name", () => {
    const c: Conversation = {
      id: "19:meet@thread.v2",
      threadProperties: { threadType: "meeting" },
      lastMessage: msg({
        id: "1",
        from: "8:other",
        imdisplayname: "Pat",
      }),
    };
    expect(conversationTitle(c)).toBe("Meeting");
  });

  it("defaults dm stub without metadata to Direct message", () => {
    expect(conversationTitle({ id: "x" })).toBe("Direct message");
  });
});

describe("messagePlainText", () => {
  it("returns plain strings unchanged", () => {
    expect(messagePlainText(" hi ")).toBe("hi");
  });

  it("strips simple HTML", () => {
    expect(messagePlainText("<p>One</p><p>Two</p>")).toBe("OneTwo");
  });

  it("handles empty", () => {
    expect(messagePlainText("")).toBe("");
  });
});

describe("isSelfMessage", () => {
  it("matches 8:skypeId form", () => {
    expect(isSelfMessage("8:abc", "abc")).toBe(true);
  });

  it("matches when MRI ends with skypeId", () => {
    expect(isSelfMessage("8:orgid:abc-123", "abc-123")).toBe(true);
  });

  it("returns false without skypeId", () => {
    expect(isSelfMessage("8:abc", undefined)).toBe(false);
  });

  it("matches self when from is a contacts URL", () => {
    expect(
      isSelfMessage(
        "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/contacts/8:orgid:abc-123",
        "orgid:abc-123",
      ),
    ).toBe(true);
  });
});

describe("formatMessageTime", () => {
  it("formats valid iso", () => {
    const s = formatMessageTime("2020-01-15T14:30:00.000Z");
    expect(s.length).toBeGreaterThan(0);
  });

  it("returns empty for invalid", () => {
    expect(formatMessageTime("not-a-date")).toBe("");
  });
});

describe("formatDetailedTimestamp", () => {
  it("formats valid iso", () => {
    const s = formatDetailedTimestamp("2020-01-15T14:30:00.000Z");
    expect(s.length).toBeGreaterThan(0);
  });

  it("returns empty for invalid", () => {
    expect(formatDetailedTimestamp("not-a-date")).toBe("");
  });
});

describe("formatDayLabel", () => {
  it("returns Today for current calendar day", () => {
    const now = new Date();
    expect(formatDayLabel(now.toISOString())).toBe("Today");
  });

  it("formats other calendar days", () => {
    const s = formatDayLabel("2019-03-15T12:00:00.000Z");
    expect(s.length).toBeGreaterThan(0);
    expect(s).not.toBe("Today");
  });
});

describe("messageTimestamp", () => {
  it("prefers originalarrivaltime", () => {
    const m = msg({
      id: "1",
      from: "x",
      originalarrivaltime: "a",
      composetime: "b",
    });
    expect(messageTimestamp(m)).toBe("a");
  });
});

describe("conversationPreview", () => {
  it("truncates long plain text", () => {
    const long = "x".repeat(100);
    const c: Conversation = {
      id: "c",
      lastMessage: msg({ id: "1", from: "f", content: long }),
    };
    expect(conversationPreview(c).endsWith("…")).toBe(true);
  });

  it("summarizes non-chat last messages", () => {
    const c: Conversation = {
      id: "c",
      lastMessage: msg({
        id: "1",
        from: "f",
        messagetype: "Event",
        content: "{}",
      }),
    };
    expect(conversationPreview(c)).toBe("Meeting or system activity");
  });

  it("uses activity metadata when content is empty system payload", () => {
    const c: Conversation = {
      id: "c",
      lastMessage: msg({
        id: "1",
        from: "8:f",
        type: "ThreadActivity",
        messagetype: "Event",
        content: "{}",
        properties: {
          activity: {
            sourceUserImDisplayName: "Pat Lee",
            activityOperationType: "addedMember",
            messagePreview: "Jordan Rivera",
          },
        },
      }),
    };
    expect(conversationPreview(c)).toBe("Pat Lee added member: Jordan Rivera");
  });

  it("combines quote and reply in preview", () => {
    const c: Conversation = {
      id: "c",
      lastMessage: msg({
        id: "1",
        from: "f",
        content:
          "Please send the report by Friday.\nThanks, I'll send it tomorrow.",
      }),
    };
    expect(conversationPreview(c)).toContain("·");
    expect(conversationPreview(c)).toContain("Thanks");
  });
});

describe("missing content", () => {
  it("does not throw when content is absent", () => {
    const m = msg({ id: "1", from: "8:x" });
    delete (m as { content?: string }).content;
    expect(parseMessageQuoteAndBody(undefined)).toEqual({
      quote: null,
      body: "",
    });
    expect(isRenderableChatMessage(m)).toBe(false);
    expect(messagePartsForDisplay(m)).toBeNull();
  });
});

describe("parseMessageQuoteAndBody", () => {
  it("splits two-line reply after a quoted ask", () => {
    const content =
      "Also would be good to get an update on Eric usage - even if they haven't given feedback\nYes, I'll pull the data";
    const { quote, body } = parseMessageQuoteAndBody(content);
    expect(quote).toContain("Eric");
    expect(body).toContain("pull the data");
  });

  it("extracts skype reply blockquote and body", () => {
    const html =
      '<blockquote itemtype="http://schema.skype.com/Reply">Prior text here</blockquote><p>New reply</p>';
    const { quote, body } = parseMessageQuoteAndBody(html);
    expect(quote).toContain("Prior text");
    expect(body).toContain("New reply");
  });

  it("parses greater-than prefixed quote lines", () => {
    const { quote, body } = parseMessageQuoteAndBody(
      "> quoted line one\n> quoted two\nMy answer here",
    );
    expect(quote).toContain("quoted line one");
    expect(body).toContain("My answer");
  });

  it("splits on blank paragraph boundary", () => {
    const { quote, body } = parseMessageQuoteAndBody(
      "First paragraph with enough length here.\n\nSecond reply paragraph.",
    );
    expect(quote).toContain("First paragraph");
    expect(body).toContain("Second reply");
  });

  it("splits on triple newline", () => {
    const { quote, body } = parseMessageQuoteAndBody(
      "Quoted block here with text.\n\n\nReply paragraph follows.",
    );
    expect(quote).toContain("Quoted block");
    expect(body).toContain("Reply paragraph");
  });

  it("preserves newlines between block elements in a blockquote", () => {
    const html =
      '<blockquote itemtype="http://schema.skype.com/Reply"><div><b>Siphesihle Thomo</b></div><div>Wanna schedule the release for tonight?</div></blockquote><p>Yes lets do it</p>';
    const { quote, body } = parseMessageQuoteAndBody(html);
    expect(quote).toContain("Siphesihle Thomo");
    expect(quote).toContain("Wanna schedule");
    expect(quote).toContain("\n");
    expect(body).toContain("Yes lets do it");
  });
});

describe("messagePartsForDisplay", () => {
  it("returns separate quote and body", () => {
    const m = msg({
      id: "1",
      from: "x",
      content:
        "Can we get an update on Eric usage even without formal feedback?\nI agree we should track that.",
    });
    const parts = messagePartsForDisplay(m);
    expect(parts?.quote).toContain("Eric");
    expect(parts?.body).toContain("track");
  });
});

describe("messageRichPartsForDisplay", () => {
  it("preserves anchors and mentions from html messages", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "rich-1",
        from: "x",
        content:
          '<div>Hello <at id="0">Dirk</at> see <a href="https://example.com/spec">spec</a></div>',
      }),
    );
    expect(parts?.quote).toBeNull();
    expect(parts?.body).toEqual([
      { kind: "text", text: "Hello " },
      { kind: "mention", text: "@Dirk" },
      { kind: "text", text: " see " },
      {
        kind: "link",
        text: "spec",
        href: "https://example.com/spec",
      },
      { kind: "text", text: "\n" },
    ]);
  });

  it("collapses Teams spacer rows between paragraphs", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "rich-2",
        from: "x",
        content:
          "<div>Hey Martha!</div><div>&nbsp;</div><div>I figured out what the issue is.</div><div>&nbsp;</div><div>Will explain the difference in our 1-on-1.</div>",
      }),
    );
    expect(parts?.body).toEqual([
      {
        kind: "text",
        text: "Hey Martha!\n\nI figured out what the issue is.\n\nWill explain the difference in our 1-on-1.\n",
      },
    ]);
  });

  it("captures message references from mention metadata", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "rich-3",
        from: "x",
        conversationId: "19:test@thread.v2",
        content: '<div>See <at data-message-id="42">that message</at></div>',
      }),
    );
    expect(parts?.body).toEqual([
      { kind: "text", text: "See " },
      {
        kind: "mention",
        text: "@that message",
        messageRef: {
          conversationId: "19:test@thread.v2",
          messageId: "42",
        },
      },
      { kind: "text", text: "\n" },
    ]);
    expect(parts?.quoteRef).toBeNull();
  });

  it("captures user mention identity from message mention metadata", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "rich-mention-user",
        from: "x",
        conversationId: "19:test@thread.v2",
        content: '<div>Hello <at id="0">Siphesihle Thomo</at></div>',
        properties: {
          mentions: [
            {
              id: "0",
              mri: "8:orgid:peer-123",
              displayName: "Siphesihle Thomo",
            },
          ],
        },
      }),
    );
    expect(parts?.body).toEqual([
      { kind: "text", text: "Hello " },
      {
        kind: "mention",
        text: "@Siphesihle Thomo",
        mentionedMri: "8:orgid:peer-123",
        mentionedDisplayName: "Siphesihle Thomo",
      },
      { kind: "text", text: "\n" },
    ]);
  });

  it("reuses cached rich parts for the same message object", () => {
    const message = msg({
      id: "rich-cache",
      from: "x",
      content: "<div>Hello again</div>",
    });

    const first = messageRichPartsForDisplay(message);
    const second = messageRichPartsForDisplay(message);

    expect(second).toBe(first);
  });

  it("merges adjacent mention fragments that target the same person", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "rich-mention-merged",
        from: "x",
        conversationId: "19:test@thread.v2",
        content: '<div><at id="0">Kane</at> <at id="1">Mooi</at></div>',
        properties: {
          mentions: [
            {
              id: "0",
              mri: "8:orgid:kane-mooi",
              displayName: "Kane Mooi",
            },
            {
              id: "1",
              mri: "8:orgid:kane-mooi",
              displayName: "Kane Mooi",
            },
          ],
        },
      }),
    );
    expect(parts?.body).toEqual([
      {
        kind: "mention",
        text: "@Kane Mooi",
        mentionedMri: "8:orgid:kane-mooi",
        mentionedDisplayName: "Kane Mooi",
      },
      { kind: "text", text: "\n" },
    ]);
  });

  it("collapses extra spacer newlines after message mentions", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "rich-4",
        from: "x",
        conversationId: "19:test@thread.v2",
        content:
          '<div><at data-message-id="42">that message</at></div><div>&nbsp;</div><div>Follow-up text</div>',
      }),
    );
    expect(parts?.body).toEqual([
      {
        kind: "mention",
        text: "@that message",
        messageRef: {
          conversationId: "19:test@thread.v2",
          messageId: "42",
        },
      },
      { kind: "text", text: "\nFollow-up text\n" },
    ]);
    expect(parts?.quoteRef).toBeNull();
  });

  it("captures quote targets from qtdMsgs metadata", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "rich-5",
        from: "x",
        conversationId: "19:test@thread.v2",
        content:
          '<blockquote itemtype="http://schema.skype.com/Reply"><div>Siphesihle Thomo</div><div>Good morning Dirk</div></blockquote><div>Will do that</div>',
        properties: {
          qtdMsgs: [{ messageId: "42", sender: "8:orgid:peer" }],
        },
      }),
    );
    expect(parts?.quoteRef).toEqual({
      conversationId: "19:test@thread.v2",
      messageId: "42",
    });
  });
});

describe("normalizePreviewText", () => {
  it("collapses whitespace and strips Play suffix", () => {
    expect(normalizePreviewText("a   b\tPlay")).toBe("a b");
  });
});

describe("isLikelySystemOrCallBlob", () => {
  it("detects flightproxy call urls", () => {
    expect(
      isLikelySystemOrCallBlob(
        "https://api.flightproxy.teams.microsoft.com/api/v2/ep/x",
      ),
    ).toBe(true);
  });

  it("detects Teams call log stub lines", () => {
    expect(
      isLikelySystemOrCallBlob(
        "Call Logs for Call 8bb441e9-6340-41d8-b7c9-b5741c16abfd",
      ),
    ).toBe(true);
  });
});

describe("textLooksLikeTeamsCallLogStub", () => {
  it("matches call log line with uuid", () => {
    expect(
      textLooksLikeTeamsCallLogStub(
        "Call Logs for Call 8bb441e9-6340-41d8-b7c9-b5741c16abfd",
      ),
    ).toBe(true);
  });

  it("does not match normal chat", () => {
    expect(textLooksLikeTeamsCallLogStub("Call you later about logs")).toBe(
      false,
    );
  });
});

describe("includeConversationInSidebar", () => {
  it("excludes conversations with no lastMessage", () => {
    const c: Conversation = { id: "19:empty@thread.v2" };
    expect(includeConversationInSidebar(c)).toBe(false);
  });

  it("includes normal chats", () => {
    const c: Conversation = {
      id: "19:ok@thread.v2",
      lastMessage: msg({ id: "1", from: "8:x", content: "Hi" }),
    };
    expect(includeConversationInSidebar(c)).toBe(true);
  });

  it("excludes special notification threads", () => {
    const c: Conversation = {
      id: "48:notifications",
      lastMessage: msg({ id: "1", from: "8:x", content: "Someone reacted" }),
    };
    expect(includeConversationInSidebar(c)).toBe(false);
  });

  it("keeps self notes threads available", () => {
    const c: Conversation = {
      id: "48:notes",
      lastMessage: msg({ id: "1", from: "8:me", content: "Scratchpad" }),
    };
    expect(includeConversationInSidebar(c)).toBe(true);
  });

  it("excludes call log stubs", () => {
    const c: Conversation = {
      id: "19:stub@thread.v2",
      lastMessage: msg({
        id: "1",
        from: "8:me",
        content: "Call Logs for Call 8bb441e9-6340-41d8-b7c9-b5741c16abfd",
      }),
    };
    expect(includeConversationInSidebar(c)).toBe(false);
  });

  it("excludes Viva / company communications team threads", () => {
    const c: Conversation = {
      id: "19:team@thread.tacv2",
      threadProperties: {
        spaceThreadTopic: "Company Communications",
        sharepointSiteUrl:
          "https://contoso.sharepoint.com/sites/CompanyCommunications2",
      },
      lastMessage: msg({
        id: "1",
        from: "8:x",
        content: "go on Edward Wilton",
      }),
    };
    expect(includeConversationInSidebar(c)).toBe(false);
  });
});

describe("filterConversationsForPipeline", () => {
  const lm = msg({ id: "1", from: "8:x", content: "focus sessions" });

  it("drops Azure Audit team space, channels by groupId, and Azure Expert MSP meetings", () => {
    const teamId = "19:teamspace@thread.tacv2";
    const team: Conversation = {
      id: teamId,
      threadProperties: {
        threadType: "space",
        productThreadType: "TeamsTeam",
        spaceThreadTopic: "Azure Audit and Specializations",
        sharepointSiteUrl:
          "https://contoso.sharepoint.com/sites/AzureAuditandSpecialisations",
        groupId: "fe9471cb-62ed-43bc-a276-45d44240d793",
      },
      lastMessage: lm,
    };
    const mspChannel: Conversation = {
      id: "19:msp@thread.tacv2",
      threadProperties: {
        threadType: "topic",
        topic: "Azure Expert MSP 2024",
        topicThreadTopic: "Azure Expert MSP 2024",
        spaceId: teamId,
        groupId: "fe9471cb-62ed-43bc-a276-45d44240d793",
      },
      lastMessage: lm,
    };
    const siblingChannel: Conversation = {
      id: "19:kube@thread.tacv2",
      threadProperties: {
        threadType: "topic",
        topic: "Kubernetes on Microsoft Azure",
        topicThreadTopic: "Kubernetes on Microsoft Azure",
        spaceId: teamId,
        groupId: "fe9471cb-62ed-43bc-a276-45d44240d793",
      },
      lastMessage: lm,
    };
    const mspMeeting: Conversation = {
      id: "19:meeting_x@thread.v2",
      threadProperties: {
        threadType: "meeting",
        topic: "Azure Expert MSP Audit 2024 | Kick-off Meeting",
      },
      lastMessage: lm,
    };
    const other: Conversation = {
      id: "19:other@thread.v2",
      lastMessage: lm,
    };
    const input = [team, mspChannel, siblingChannel, mspMeeting, other];
    const out = filterConversationsForPipeline(input);
    expect(out).toEqual([other]);
  });

  it("keeps unrelated Azure meetings", () => {
    const meeting: Conversation = {
      id: "19:meeting_q@thread.v2",
      threadProperties: {
        threadType: "meeting",
        topic: "Azure Quotas - Design/Scope",
      },
      lastMessage: lm,
    };
    expect(filterConversationsForPipeline([meeting])).toEqual([meeting]);
  });

  it("drops topic and spaceId conversations immediately", () => {
    const topicOnly: Conversation = {
      id: "19:topiconly@thread.tacv2",
      threadProperties: { threadType: "topic", topic: "General" },
      lastMessage: lm,
    };
    const hasSpaceId: Conversation = {
      id: "19:spacechild@thread.tacv2",
      threadProperties: { spaceId: "19:parent@thread.tacv2", topic: "Any" },
      lastMessage: lm,
    };
    const normalDm: Conversation = {
      id: "19:normaldm@unq.gbl.spaces",
      threadProperties: {
        threadType: "chat",
        productThreadType: "OneToOneChat",
      },
      lastMessage: lm,
    };
    expect(
      filterConversationsForPipeline([topicOnly, hasSpaceId, normalDm]),
    ).toEqual([normalDm]);
  });

  it("drops chats whose last message text contains Azure Expert MSP", () => {
    const mspInMessage: Conversation = {
      id: "19:oddshape@unq.gbl.spaces",
      lastMessage: msg({
        id: "1",
        from: "8:x",
        content:
          "Azure Expert MSP 2024 for this week's focus sessions, please can you all bring your evidence",
      }),
    };
    const normalDm: Conversation = {
      id: "19:normaldm@unq.gbl.spaces",
      lastMessage: msg({ id: "2", from: "8:y", content: "Hey, quick update." }),
    };
    const out = filterConversationsForPipeline([mspInMessage, normalDm]);
    expect(out).toEqual([normalDm]);
  });

  it("treats Azure audit team as root when threadType is missing but product is TeamsTeam", () => {
    const teamId = "19:azureteam@thread.tacv2";
    const team: Conversation = {
      id: teamId,
      threadProperties: {
        productThreadType: "TeamsTeam",
        sharepointSiteUrl:
          "https://contoso.sharepoint.com/sites/AzureAuditandSpecialisations",
        groupId: "aaa",
      },
      lastMessage: lm,
    };
    const channel: Conversation = {
      id: "19:ch@thread.tacv2",
      threadProperties: {
        threadType: "topic",
        topic: "Other channel",
        spaceId: teamId,
        groupId: "aaa",
      },
      lastMessage: lm,
    };
    expect(filterConversationsForPipeline([team, channel])).toEqual([]);
  });
});

describe("isCompanyCommunicationsSidebarThread", () => {
  it("is true for spaceThreadTopic Company Communications", () => {
    const c: Conversation = {
      id: "19:t@thread.tacv2",
      threadProperties: { spaceThreadTopic: "Company Communications" },
    };
    expect(isCompanyCommunicationsSidebarThread(c)).toBe(true);
  });

  it("is true for Viva Company Announcements topic channel", () => {
    const c: Conversation = {
      id: "19:topic@thread.tacv2",
      threadProperties: {
        topic: "Viva Company Announcements",
        topicThreadTopic: "Viva Company Announcements",
      },
    };
    expect(isCompanyCommunicationsSidebarThread(c)).toBe(true);
  });

  it("is true when topics JSON names Viva Company Announcements", () => {
    const c: Conversation = {
      id: "19:t@thread.tacv2",
      threadProperties: {
        topics:
          '[{"name":"Viva Company Announcements","id":"19:x@thread.tacv2"}]',
      },
    };
    expect(isCompanyCommunicationsSidebarThread(c)).toBe(true);
  });

  it("is false for unrelated team chat", () => {
    const c: Conversation = {
      id: "19:g@thread.tacv2",
      threadProperties: { topic: "Engineering" },
    };
    expect(isCompanyCommunicationsSidebarThread(c)).toBe(false);
  });

  it("is false when threadProperties absent", () => {
    const c: Conversation = { id: "19:x@thread.v2" };
    expect(isCompanyCommunicationsSidebarThread(c)).toBe(false);
  });

  it("is true when title-like properties carry Company Communications", () => {
    const c: Conversation = {
      id: "19:prop@thread.tacv2",
      properties: { msTeamsThreadName: "Company Communications" },
    };
    expect(isCompanyCommunicationsSidebarThread(c)).toBe(true);
  });
});

describe("isCallLogsStubConversation", () => {
  it("is true when last message is a call log stub", () => {
    const c: Conversation = {
      id: "19:stub@thread.v2",
      lastMessage: msg({
        id: "1",
        from: "8:me",
        content: "Call Logs for Call 8bb441e9-6340-41d8-b7c9-b5741c16abfd",
      }),
    };
    expect(isCallLogsStubConversation(c)).toBe(true);
  });

  it("is false for regular last message", () => {
    const c: Conversation = {
      id: "19:ok@thread.v2",
      lastMessage: msg({ id: "1", from: "8:you", content: "Hello there" }),
    };
    expect(isCallLogsStubConversation(c)).toBe(false);
  });
});

describe("isRenderableChatMessage", () => {
  it("accepts Event messages with readable text", () => {
    expect(
      isRenderableChatMessage(
        msg({ id: "1", from: "x", messagetype: "Event", content: "hi" }),
      ),
    ).toBe(true);
  });

  it("rejects long MRI dumps", () => {
    const a = "8:orgid:aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee";
    const b = "8:orgid:bbbbbbbb-cccc-dddd-eeee-ffffffffffff";
    expect(
      isRenderableChatMessage(
        msg({ id: "1", from: "x", content: `${a} Name1 ${b} Name2` }),
      ),
    ).toBe(false);
  });

  it("accepts activity messages with structured activity metadata", () => {
    expect(
      isRenderableChatMessage(
        msg({
          id: "1",
          from: "8:x",
          type: "ThreadActivity",
          messagetype: "Event",
          content: "{}",
          properties: {
            activity: {
              sourceUserImDisplayName: "Pat Lee",
              activityOperationType: "addedMember",
              messagePreview: "Jordan Rivera",
            },
          },
        }),
      ),
    ).toBe(true);
  });
});

describe("messageBodyForDisplay", () => {
  it("returns null for filtered messages", () => {
    expect(
      messageBodyForDisplay(
        msg({ id: "1", from: "x", messagetype: "Event", content: "x" }),
      ),
    ).toBe("x");
  });

  it("returns a readable summary for activity metadata", () => {
    expect(
      messageBodyForDisplay(
        msg({
          id: "1",
          from: "8:x",
          type: "ThreadActivity",
          messagetype: "Event",
          content: "{}",
          properties: {
            activity: {
              sourceUserImDisplayName: "Pat Lee",
              activityOperationType: "addedMember",
              messagePreview: "Jordan Rivera",
            },
          },
        }),
      ),
    ).toBe("Pat Lee added member: Jordan Rivera");
  });

  it("joins quote and reply with blank line", () => {
    const out = messageBodyForDisplay(
      msg({
        id: "1",
        from: "x",
        content: "Need the numbers by Tuesday.\nThanks, will send Monday.",
      }),
    );
    expect(out).toContain("Tuesday");
    expect(out).toContain("Thanks");
    expect(out).toMatch(/\n\n/);
  });
});

describe("sortConversationsByActivity", () => {
  it("orders by last message time descending", () => {
    const older: Conversation = {
      id: "o",
      lastMessage: msg({
        id: "1",
        from: "a",
        originalarrivaltime: "2020-01-01T00:00:00.000Z",
      }),
    };
    const newer: Conversation = {
      id: "n",
      lastMessage: msg({
        id: "2",
        from: "b",
        originalarrivaltime: "2021-01-01T00:00:00.000Z",
      }),
    };
    const sorted = sortConversationsByActivity([older, newer]);
    expect(sorted[0].id).toBe("n");
  });
});

describe("gapBetweenMessages", () => {
  it("computes positive delta in ms", () => {
    const a = msg({
      id: "1",
      from: "x",
      originalarrivaltime: "2020-01-01T00:00:00.000Z",
    });
    const b = msg({
      id: "2",
      from: "x",
      originalarrivaltime: "2020-01-01T00:01:00.000Z",
    });
    expect(gapBetweenMessages(a, b)).toBe(60_000);
  });
});

// ── Attachment rendering tests (Siphesihle Thomo chat patterns) ──

describe("attachment rendering with Siphesihle Thomo chat data", () => {
  it("extracts image attachment from Picture.1 URIObject", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "img-1",
        from: "8:orgid:siphesihle-uuid",
        imdisplayname: "Siphesihle Thomo",
        messagetype: "RichText/UriObject",
        contenttype: "RichText/UriObject",
        content:
          '<URIObject type="Picture.1" uri="https://api.asm.skype.com/v1/objects/0-sa-d1-img001" url_thumbnail="https://api.asm.skype.com/v1/objects/0-sa-d1-img001/views/imgt1"><Title>Title: screenshot.png</Title><Description>Description: screenshot.png</Description><OriginalName v="screenshot.png"/><FileSize v="145920"/><a href="https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-sa-d1-img001">https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-sa-d1-img001</a></URIObject>',
      }),
    );

    expect(parts).not.toBeNull();
    expect(parts?.attachments).toHaveLength(1);
    expect(parts?.attachments[0]).toEqual({
      kind: "image",
      objectUrl: "https://api.asm.skype.com/v1/objects/0-sa-d1-img001",
      openUrl:
        "https://login.skype.com/login/sso?go=webclient.xmm&docid=0-sa-d1-img001",
      thumbnailUrl:
        "https://api.asm.skype.com/v1/objects/0-sa-d1-img001/views/imgt1",
      title: "screenshot.png",
      fileName: "screenshot.png",
      fileExtension: "png",
      fileSize: 145920,
    });
    expect(parts?.body).toEqual([]);
  });

  it("extracts attachment from URIObject wrapped in div", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "wrapped-1",
        from: "8:orgid:siphesihle-uuid",
        messagetype: "RichText/Media_GenericFile",
        contenttype: "RichText/Media_GenericFile",
        content:
          '<div><URIObject type="File.1" uri="https://api.asm.skype.com/v1/objects/0-sa-d2-file"><Title>Title: report.xlsx</Title><OriginalName v="report.xlsx"/><FileSize v="34567"/><a href="https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-sa-d2-file">link</a></URIObject></div>',
      }),
    );

    expect(parts).not.toBeNull();
    expect(parts?.attachments).toHaveLength(1);
    expect(parts?.attachments[0]?.kind).toBe("file");
    expect(parts?.attachments[0]?.title).toBe("report.xlsx");
    expect(parts?.attachments[0]?.fileExtension).toBe("xlsx");
    expect(parts?.attachments[0]?.fileSize).toBe(34567);
    // Should have empty body since it's pure attachment markup
    expect(parts?.body).toEqual([]);
  });

  it("renders attachment when amsreferences is set but messagetype is Text", () => {
    const m = msg({
      id: "ams-fallback",
      from: "8:orgid:siphesihle-uuid",
      messagetype: "Text",
      contenttype: "text",
      content:
        '<URIObject type="File.1" uri="https://api.asm.skype.com/v1/objects/0-sa-d3"><Title>Title: notes.txt</Title><OriginalName v="notes.txt"/><a href="https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-sa-d3">link</a></URIObject>',
    }) as Message & { amsreferences: string[] };
    m.amsreferences = ["0-sa-d3"];

    const parts = messageRichPartsForDisplay(m);
    expect(parts).not.toBeNull();
    expect(parts?.attachments).toHaveLength(1);
    expect(parts?.attachments[0]?.title).toBe("notes.txt");
  });

  it("extracts file extension for various file types", () => {
    const makeAttMsg = (name: string) =>
      msg({
        id: `ext-${name}`,
        from: "8:orgid:siphesihle-uuid",
        messagetype: "RichText/Media_GenericFile",
        contenttype: "RichText/Media_GenericFile",
        content: `<URIObject type="File.1" uri="https://api.asm.skype.com/v1/objects/0-${name}"><OriginalName v="${name}"/><a href="https://example.com">link</a></URIObject>`,
      });

    const pdfParts = messageRichPartsForDisplay(makeAttMsg("doc.pdf"));
    expect(pdfParts?.attachments[0]?.fileExtension).toBe("pdf");

    const docxParts = messageRichPartsForDisplay(makeAttMsg("report.docx"));
    expect(docxParts?.attachments[0]?.fileExtension).toBe("docx");

    const noExtParts = messageRichPartsForDisplay(makeAttMsg("README"));
    expect(noExtParts?.attachments[0]?.fileExtension).toBeUndefined();
  });

  it("renders message with both text body and inline attachment", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "mixed-1",
        from: "8:orgid:siphesihle-uuid",
        messagetype: "RichText/Media_GenericFile",
        contenttype: "RichText/Media_GenericFile",
        content:
          'Here is the file you asked for <URIObject type="File.1" uri="https://api.asm.skype.com/v1/objects/0-mixed"><OriginalName v="data.csv"/><a href="https://example.com">link</a></URIObject>',
      }),
    );

    expect(parts).not.toBeNull();
    expect(parts?.attachments).toHaveLength(1);
    expect(parts?.attachments[0]?.title).toBe("data.csv");
    // Body should contain the text before the URIObject
    const bodyText = parts?.body.map((p) => p.text).join("") ?? "";
    expect(bodyText).toContain("Here is the file");
  });

  it("handles quote reply with attachment from Siphesihle", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "quote-att-1",
        from: "8:orgid:dirk-uuid",
        conversationId: "19:test@thread.v2",
        messagetype: "RichText/Media_GenericFile",
        contenttype: "RichText/Media_GenericFile",
        content:
          '<blockquote itemtype="http://schema.skype.com/Reply"><div><b>Siphesihle Thomo</b></div><div>Can you send me the updated spec?</div></blockquote><URIObject type="File.1" uri="https://api.asm.skype.com/v1/objects/0-spec"><OriginalName v="spec-v2.pdf"/><FileSize v="512000"/><a href="https://example.com/spec">link</a></URIObject>',
        properties: {
          qtdMsgs: [{ messageId: "99", sender: "8:orgid:siphesihle-uuid" }],
        },
      }),
    );

    expect(parts).not.toBeNull();
    expect(parts?.quote).not.toBeNull();
    const quoteText = parts?.quote?.map((p) => p.text).join("") ?? "";
    expect(quoteText).toContain("Siphesihle Thomo");
    expect(quoteText).toContain("updated spec");
    expect(parts?.attachments).toHaveLength(1);
    expect(parts?.attachments[0]?.title).toBe("spec-v2.pdf");
    expect(parts?.quoteRef).toEqual({
      conversationId: "19:test@thread.v2",
      messageId: "99",
    });
  });
});

// ── Deleted messages ──

describe("deleted message handling", () => {
  it("renders deleted message as a placeholder", () => {
    const m = msg({
      id: "del-1",
      from: "8:orgid:siphesihle-uuid",
      content: "",
      deleted: true,
      properties: { deletetime: 1711446123456 },
    });

    expect(isRenderableChatMessage(m)).toBe(true);
    const parts = messageRichPartsForDisplay(m);
    expect(parts).not.toBeNull();
    expect(parts?.body).toEqual([
      { kind: "text", text: "This message has been deleted." },
    ]);
    expect(parts?.attachments).toEqual([]);
    expect(parts?.quote).toBeNull();
  });

  it("messagePartsForDisplay returns deleted placeholder text", () => {
    const m = msg({
      id: "del-2",
      from: "8:orgid:siphesihle-uuid",
      content: "",
      deleted: true,
      properties: { deletetime: 1711446123456 },
    });

    const parts = messagePartsForDisplay(m);
    expect(parts).not.toBeNull();
    expect(parts?.body).toBe("This message has been deleted.");
    expect(parts?.quote).toBeNull();
  });

  it("hard-deleted messages also render as placeholders", () => {
    const m = msg({
      id: "del-3",
      from: "8:orgid:siphesihle-uuid",
      content: "",
      deleted: true,
      properties: { hardDeleteTime: 1711446200000, hardDeleteReason: "user" },
    });

    expect(isRenderableChatMessage(m)).toBe(true);
    const parts = messageRichPartsForDisplay(m);
    expect(parts?.body[0]?.text).toBe("This message has been deleted.");
  });
});

// ── Edited messages ──

describe("isEditedMessage", () => {
  it("detects edited message from edittime property", () => {
    const m = msg({
      id: "edit-1",
      from: "8:orgid:siphesihle-uuid",
      content: "Updated text",
      properties: { edittime: "1711446123456" },
    });
    expect(isEditedMessage(m)).toBe(true);
  });

  it("returns false when edittime is missing", () => {
    const m = msg({
      id: "edit-2",
      from: "8:orgid:siphesihle-uuid",
      content: "Original text",
    });
    expect(isEditedMessage(m)).toBe(false);
  });

  it("returns false when edittime is zero", () => {
    const m = msg({
      id: "edit-3",
      from: "8:orgid:siphesihle-uuid",
      content: "Text",
      properties: { edittime: "0" },
    });
    expect(isEditedMessage(m)).toBe(false);
  });

  it("returns false when edittime is numeric zero", () => {
    const m = msg({
      id: "edit-4",
      from: "8:orgid:siphesihle-uuid",
      content: "Text",
      properties: { edittime: 0 },
    });
    expect(isEditedMessage(m)).toBe(false);
  });
});

// ── Consumption Horizon & Read Receipts ──

describe("parseConsumptionHorizon", () => {
  it("parses valid horizon string", () => {
    const result = parseConsumptionHorizon(
      "1711446123456;1711446123456;2132503743217489806",
    );
    expect(result).toEqual({
      sequenceId: 1711446123456,
      timestamp: 1711446123456,
      messageId: "2132503743217489806",
    });
  });

  it("returns null for empty string", () => {
    expect(parseConsumptionHorizon("")).toBeNull();
  });

  it("returns null for undefined", () => {
    expect(parseConsumptionHorizon(undefined)).toBeNull();
  });

  it("returns null for malformed string with fewer than 3 parts", () => {
    expect(parseConsumptionHorizon("123;456")).toBeNull();
  });

  it("returns null for non-numeric sequence id", () => {
    expect(parseConsumptionHorizon("abc;123;2132503743217489806")).toBeNull();
  });

  it("preserves message id with semicolons", () => {
    const result = parseConsumptionHorizon("100;200;3572204487355094743;extra");
    expect(result?.messageId).toBe("3572204487355094743;extra");
  });
});

describe("messageReadStatus", () => {
  it("returns read when peer horizon exceeds message sequence", () => {
    const m = msg({
      id: "100",
      from: "8:orgid:self",
      sequenceId: 100,
    } as Partial<Message> & Pick<Message, "id" | "from">);
    (m as Message & { sequenceId: number }).sequenceId = 100;

    const status = messageReadStatus(m, [
      { sequenceId: 150, timestamp: 0, messageId: "150" },
    ]);
    expect(status).toBe("read");
  });

  it("returns delivered when peer horizon is below message sequence", () => {
    const m = msg({
      id: "200",
      from: "8:orgid:self",
    });
    (m as Message & { sequenceId: number }).sequenceId = 200;

    const status = messageReadStatus(m, [
      { sequenceId: 100, timestamp: 0, messageId: "100" },
    ]);
    expect(status).toBe("delivered");
  });

  it("returns sent when no peer horizons are available", () => {
    const m = msg({
      id: "300",
      from: "8:orgid:self",
    });
    (m as Message & { sequenceId: number }).sequenceId = 300;

    const status = messageReadStatus(m, []);
    expect(status).toBe("sent");
  });

  it("returns read when peer horizon equals message sequence", () => {
    const m = msg({
      id: "50",
      from: "8:orgid:self",
    });
    (m as Message & { sequenceId: number }).sequenceId = 50;

    const status = messageReadStatus(m, [
      { sequenceId: 50, timestamp: 0, messageId: "50" },
    ]);
    expect(status).toBe("read");
  });

  it("returns the newest matching read timestamp", () => {
    const m = msg({
      id: "100",
      from: "8:orgid:self",
    });
    (m as Message & { sequenceId: number }).sequenceId = 100;

    expect(
      messageReadTimestamp(m, [
        {
          sequenceId: 100,
          timestamp: 1_711_446_123_456,
          messageId: "100",
        },
        {
          sequenceId: 140,
          timestamp: 1_711_446_223_456,
          messageId: "140",
        },
      ]),
    ).toBe("2024-03-26T09:43:43.456Z");
  });

  it("returns empty when the message has not been read", () => {
    const m = msg({
      id: "200",
      from: "8:orgid:self",
    });
    (m as Message & { sequenceId: number }).sequenceId = 200;

    expect(
      messageReadTimestamp(m, [
        {
          sequenceId: 100,
          timestamp: 1_711_446_123_456,
          messageId: "100",
        },
      ]),
    ).toBe("");
  });
});

// ── Siphesihle Thomo specific conversation patterns ──

describe("Siphesihle Thomo conversation patterns", () => {
  it("renders rich quote with Siphesihle as author and preserves formatting", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "sip-quote-1",
        from: "8:orgid:dirk-uuid",
        conversationId: "19:siphesihle-dm@unq.gbl.spaces",
        content:
          '<blockquote itemtype="http://schema.skype.com/Reply"><div><b>Siphesihle Thomo</b></div><div>Wanna schedule the release for tonight?</div></blockquote><p>Yes lets do it</p>',
        properties: {
          qtdMsgs: [{ messageId: "42", sender: "8:orgid:siphesihle-uuid" }],
        },
      }),
    );

    expect(parts?.quote).not.toBeNull();
    const quoteText = parts?.quote?.map((p) => p.text).join("") ?? "";
    expect(quoteText).toContain("Siphesihle Thomo");
    expect(quoteText).toContain("Wanna schedule");
    expect(quoteText).toContain("\n");
    expect(parts?.quoteRef).toEqual({
      conversationId: "19:siphesihle-dm@unq.gbl.spaces",
      messageId: "42",
    });
    const bodyText = parts?.body.map((p) => p.text).join("") ?? "";
    expect(bodyText).toContain("Yes lets do it");
  });

  it("renders mention of Siphesihle with MRI from metadata", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "sip-mention-1",
        from: "8:orgid:dirk-uuid",
        conversationId: "19:group@thread.v2",
        content:
          '<div>Hey <at id="0">Siphesihle Thomo</at> check this out</div>',
        properties: {
          mentions: [
            {
              id: "0",
              mri: "8:orgid:siphesihle-uuid",
              displayName: "Siphesihle Thomo",
            },
          ],
        },
      }),
    );

    expect(parts?.body).toEqual([
      { kind: "text", text: "Hey " },
      {
        kind: "mention",
        text: "@Siphesihle Thomo",
        mentionedMri: "8:orgid:siphesihle-uuid",
        mentionedDisplayName: "Siphesihle Thomo",
      },
      { kind: "text", text: " check this out\n" },
    ]);
  });

  it("handles Siphesihle sending a message with bold + link formatting", () => {
    const parts = messageRichPartsForDisplay(
      msg({
        id: "sip-rich-1",
        from: "8:orgid:siphesihle-uuid",
        imdisplayname: "Siphesihle Thomo",
        content:
          '<div><b>Important:</b> See <a href="https://dev.azure.com/project/release">the release pipeline</a></div>',
      }),
    );

    expect(parts?.body).toEqual([
      { kind: "text", text: "Important:", bold: true },
      { kind: "text", text: " See " },
      {
        kind: "link",
        text: "the release pipeline",
        href: "https://dev.azure.com/project/release",
      },
      { kind: "text", text: "\n" },
    ]);
  });

  it("handles Siphesihle sending a deleted message that previously had content", () => {
    const m = msg({
      id: "sip-del-1",
      from: "8:orgid:siphesihle-uuid",
      imdisplayname: "Siphesihle Thomo",
      content: "",
      deleted: true,
      properties: { deletetime: 1711446200000 },
    });

    const parts = messageRichPartsForDisplay(m);
    expect(parts).not.toBeNull();
    expect(parts?.body).toEqual([
      { kind: "text", text: "This message has been deleted." },
    ]);
  });

  it("handles Siphesihle sending an edited message", () => {
    const m = msg({
      id: "sip-edit-1",
      from: "8:orgid:siphesihle-uuid",
      imdisplayname: "Siphesihle Thomo",
      content: "Updated: The meeting is at 3pm not 2pm",
      properties: { edittime: "1711446300000" },
    });

    expect(isEditedMessage(m)).toBe(true);
    const parts = messageRichPartsForDisplay(m);
    expect(parts).not.toBeNull();
    const bodyText = parts?.body.map((p) => p.text).join("") ?? "";
    expect(bodyText).toContain("meeting is at 3pm");
  });

  it("renders conversation with Siphesihle as DM correctly", () => {
    const c: Conversation = {
      id: "19:siphesihle-dm_dirk@unq.gbl.spaces",
      threadProperties: { membercount: "2" },
      members: [
        {
          id: "8:orgid:dirk-uuid",
          role: "User",
          isMri: true,
          displayName: "Dirk Beukes",
        },
        {
          id: "8:orgid:siphesihle-uuid",
          role: "User",
          isMri: true,
          displayName: "Siphesihle Thomo",
        },
      ],
      lastMessage: msg({
        id: "last-1",
        from: "8:orgid:dirk-uuid",
        imdisplayname: "Dirk Beukes",
        content: "Sounds good, let me know when ready",
      }),
    };

    expect(conversationChatKind(c)).toBe("dm");
    expect(conversationTitle(c, "orgid:dirk-uuid")).toBe("Siphesihle Thomo");
  });

  it("renders consumption horizon from Siphesihle DM", () => {
    const horizon = parseConsumptionHorizon(
      "1711446123456;1711446123456;2132503743217489806",
    );
    expect(horizon).not.toBeNull();
    expect(horizon?.messageId).toBe("2132503743217489806");
    if (!horizon) {
      throw new Error("Expected horizon");
    }

    // Simulate self message read by Siphesihle
    const m = msg({
      id: "read-msg-1",
      from: "8:orgid:dirk-uuid",
    });
    (m as Message & { sequenceId: number }).sequenceId = 1711446123400;

    const status = messageReadStatus(m, [horizon]);
    expect(status).toBe("read");
  });

  it("handles multi-paragraph message from Siphesihle without false quote split", () => {
    // This message should NOT be split into quote + body because
    // the first paragraph doesn't look like a quoted reply
    const parts = messageRichPartsForDisplay(
      msg({
        id: "sip-para-1",
        from: "8:orgid:siphesihle-uuid",
        imdisplayname: "Siphesihle Thomo",
        content:
          "<div>Hi Dirk</div><div>&nbsp;</div><div>I wanted to let you know the deployment went well.</div><div>&nbsp;</div><div>All services are running smoothly.</div>",
      }),
    );

    expect(parts?.quote).toBeNull();
    const bodyText = parts?.body.map((p) => p.text).join("") ?? "";
    expect(bodyText).toContain("Hi Dirk");
    expect(bodyText).toContain("deployment went well");
    expect(bodyText).toContain("All services are running smoothly");
  });

  it("parses real Teams URIObject where OriginalName wraps the anchor tag", () => {
    // This is the exact real-world markup from Teams where <OriginalName> is
    // not self-closing and the <a> tag is nested inside it
    const realContent =
      '<URIObject type="File.1" uri="https://eu-api.asm.skype.com/v1/objects/0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b" url_thumbnail="https://eu-api.asm.skype.com/v1/objects/0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b/views/thumbnail"><Title>Title: three-body.html</Title><Description>Description: three-body.html</Description><OriginalName v="three-body.html"><a href="https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b" title="https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b" target="_blank" rel="noreferrer noopener">https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b</a></OriginalName></URIObject>';

    const parts = messageRichPartsForDisplay(
      msg({
        id: "real-att-1",
        from: "8:orgid:dirk-uuid",
        messagetype: "RichText/Media_GenericFile",
        contenttype: "RichText/Media_GenericFile",
        content: realContent,
      }),
    );

    expect(parts).not.toBeNull();
    expect(parts?.attachments).toHaveLength(1);
    expect(parts?.attachments[0]?.kind).toBe("file");
    expect(parts?.attachments[0]?.title).toBe("three-body.html");
    expect(parts?.attachments[0]?.fileName).toBe("three-body.html");
    expect(parts?.attachments[0]?.fileExtension).toBe("html");
    expect(parts?.attachments[0]?.objectUrl).toBe(
      "https://eu-api.asm.skype.com/v1/objects/0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b",
    );
    expect(parts?.attachments[0]?.openUrl).toBe(
      "https://login.skype.com/login/sso?go=webclient.xmm&docid=0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b",
    );
    expect(parts?.attachments[0]?.thumbnailUrl).toBe(
      "https://eu-api.asm.skype.com/v1/objects/0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b/views/thumbnail",
    );
    // Pure attachment markup → body should be empty
    expect(parts?.body).toEqual([]);
  });

  it("renders short text messages like Ooof without filtering them out", () => {
    // Real API data: messagetype=RichText/Html, content=<p>Ooof</p>
    const parts = messageRichPartsForDisplay(
      msg({
        id: "ooof-1",
        from: "8:orgid:dirk-uuid",
        imdisplayname: "Dirk Beukes",
        messagetype: "RichText/Html",
        contenttype: "Text",
        content: "<p>Ooof</p>",
      }),
    );

    expect(parts).not.toBeNull();
    const bodyText = parts?.body.map((p) => p.text).join("") ?? "";
    expect(bodyText.trim()).toBe("Ooof");
  });

  it("renders SharePoint file card from properties.files when content is empty", () => {
    // Real API pattern: content="" with file metadata in properties.files
    const filesJson = JSON.stringify([
      {
        itemid: "edf1f7ce-7fd3-4f8d-9cb6-e18f78801593",
        fileName: "three-body.html",
        fileType: "html",
        fileInfo: {
          fileUrl:
            "https://backupdirect-my.sharepoint.com/personal/dirk_beukes_clouddirect_net/Documents/Microsoft%20Teams%20Chat%20Files/three-body.html",
          siteUrl:
            "https://backupdirect-my.sharepoint.com/personal/dirk_beukes_clouddirect_net/",
          shareUrl:
            "https://backupdirect-my.sharepoint.com/:u:/g/personal/dirk_beukes_clouddirect_net/share123",
        },
        "@type": "http://schema.skype.com/File",
        objectUrl:
          "https://backupdirect-my.sharepoint.com/personal/dirk_beukes_clouddirect_net/Documents/Microsoft%20Teams%20Chat%20Files/three-body.html",
        title: "three-body.html",
        state: "active",
      },
    ]);

    const parts = messageRichPartsForDisplay(
      msg({
        id: "file-card-1",
        from: "8:orgid:dirk-uuid",
        imdisplayname: "Dirk Beukes",
        messagetype: "RichText/Html",
        contenttype: "Text",
        content: "",
        properties: {
          files: filesJson,
          formatVariant: "TEAMS",
        },
      }),
    );

    expect(parts).not.toBeNull();
    expect(parts?.attachments).toHaveLength(1);
    expect(parts?.attachments[0]?.kind).toBe("file");
    expect(parts?.attachments[0]?.title).toBe("three-body.html");
    expect(parts?.attachments[0]?.fileName).toBe("three-body.html");
    expect(parts?.attachments[0]?.fileExtension).toBe("html");
    expect(parts?.attachments[0]?.openUrl).toBe(
      "https://backupdirect-my.sharepoint.com/:u:/g/personal/dirk_beukes_clouddirect_net/share123",
    );
    // Body should be empty since it's a pure file card
    expect(parts?.body).toEqual([]);
  });

  it("marks SharePoint file card messages as renderable", () => {
    const filesJson = JSON.stringify([
      {
        fileName: "report.pdf",
        title: "report.pdf",
        fileInfo: {
          shareUrl: "https://example.com/share/report.pdf",
        },
        objectUrl: "https://example.com/report.pdf",
      },
    ]);

    const m = msg({
      id: "file-card-renderable",
      from: "8:orgid:dirk-uuid",
      messagetype: "RichText/Html",
      contenttype: "Text",
      content: "",
      properties: { files: filesJson },
    });

    expect(isRenderableChatMessage(m)).toBe(true);
  });

  it("renders SharePoint screenshot as image with AMS preview URL", () => {
    // Real API data: screenshots uploaded to SharePoint have filePreview.previewUrl
    const filesJson = JSON.stringify([
      {
        itemid: "68950f5f-bb65-4147-88c4-c9c6b913184c",
        fileName: "Screenshot 2026-03-24 at 3.24.25 PM.png",
        fileType: "png",
        fileInfo: {
          fileUrl:
            "https://backupdirect-my.sharepoint.com/personal/dirk_beukes_clouddirect_net/Documents/Microsoft%20Teams%20Chat%20Files/Screenshot%202026-03-24%20at%203.24.25%E2%80%AFPM.png",
          shareUrl:
            "https://backupdirect-my.sharepoint.com/:i:/g/personal/dirk_beukes_clouddirect_net/IQBfD5VoZbtHQYjEyca5ExhMAWta1b7o4Ii17FIag3a06-g",
        },
        "@type": "http://schema.skype.com/File",
        objectUrl:
          "https://backupdirect-my.sharepoint.com/personal/dirk_beukes_clouddirect_net/Documents/Microsoft%20Teams%20Chat%20Files/Screenshot%202026-03-24%20at%203.24.25%E2%80%AFPM.png",
        title: "Screenshot 2026-03-24 at 3.24.25 PM.png",
        filePreview: {
          previewUrl:
            "https://eu-api.asm.skype.com/v1/objects/0-weu-d16-f6aa5e6f148edce088a8480476d8568f/views/imgo",
          previewHeight: 1802,
          previewWidth: 2806,
        },
      },
    ]);

    const parts = messageRichPartsForDisplay(
      msg({
        id: "screenshot-card",
        from: "8:orgid:dirk-uuid",
        imdisplayname: "Dirk Beukes",
        messagetype: "RichText/Html",
        contenttype: "Text",
        content: "<p>Nice</p>",
        properties: { files: filesJson },
      }),
    );

    expect(parts).not.toBeNull();
    expect(parts?.attachments).toHaveLength(1);
    const att = parts?.attachments[0];
    expect(att?.kind).toBe("image");
    expect(att?.fileExtension).toBe("png");
    expect(att?.thumbnailUrl).toBe(
      "https://eu-api.asm.skype.com/v1/objects/0-weu-d16-f6aa5e6f148edce088a8480476d8568f/views/imgo",
    );
    expect(att?.openUrl).toBe(
      "https://backupdirect-my.sharepoint.com/:i:/g/personal/dirk_beukes_clouddirect_net/IQBfD5VoZbtHQYjEyca5ExhMAWta1b7o4Ii17FIag3a06-g",
    );
    // Body should contain the "Nice" text
    const bodyText = parts?.body.map((p) => p.text).join("") ?? "";
    expect(bodyText.trim()).toBe("Nice");
  });

  it("renders message with text body before a URIObject attachment", () => {
    const realContent =
      'It\'s just this \n<URIObject type="File.1" uri="https://eu-api.asm.skype.com/v1/objects/0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b" url_thumbnail="https://eu-api.asm.skype.com/v1/objects/0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b/views/thumbnail"><Title>Title: three-body.html</Title><Description>Description: three-body.html</Description><OriginalName v="three-body.html"><a href="https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b" title="https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b" target="_blank" rel="noreferrer noopener">https://login.skype.com/login/sso?go=webclient.xmm&amp;docid=0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b</a></OriginalName></URIObject>';

    const parts = messageRichPartsForDisplay(
      msg({
        id: "mixed-real-1",
        from: "8:orgid:dirk-uuid",
        messagetype: "RichText/Media_GenericFile",
        contenttype: "RichText/Media_GenericFile",
        content: realContent,
      }),
    );

    expect(parts).not.toBeNull();
    expect(parts?.attachments).toHaveLength(1);
    expect(parts?.attachments[0]?.title).toBe("three-body.html");
    // Body should contain just the text before the URIObject, not the
    // URIObject element's internal text content
    const bodyText = parts?.body.map((p) => p.text).join("") ?? "";
    expect(bodyText).toContain("It's just this");
    // Should NOT contain the URI object internals
    expect(bodyText).not.toContain("Title: three-body.html");
    expect(bodyText).not.toContain("Description:");
  });

  it("renders code block message with pasted URIObject XML (Siphesihle 4:01 PM)", () => {
    // Real API data: Siphesihle pasted URIObject XML as a code block
    const realContent =
      '<p>It\'s just this&nbsp;</p>\r\n<p itemtype="http://schema.skype.com/CodeBlockEditor" id="x_codeBlockEditor-a70b39d9">\r\n&nbsp;</p>\r\n<pre class="language-json" itemid="codeBlockEditor-a70b39d9"><code>&lt;URIObject&nbsp;type=<span class="hljs-string">&quot;File.1&quot;</span>&nbsp;uri=<span class="hljs-string">&quot;https://eu-api.asm.skype.com/v1/objects/0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b&quot;</span>&nbsp;url_thumbnail=<span class="hljs-string">&quot;https://eu-api.asm.skype.com/v1/objects/0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b/views/thumbnail&quot;</span>&gt;&lt;Title&gt;Title:&nbsp;three-body.html&lt;/Title&gt;&lt;Description&gt;Description:&nbsp;three-body.html&lt;/Description&gt;&lt;OriginalName&nbsp;v=<span class="hljs-string">&quot;three-body.html&quot;</span>&gt;&lt;a&nbsp;href=<span class="hljs-string">&quot;https://login.skype.com/login/sso?go=webclient.xmm&amp;amp;docid=0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b&quot;</span>&gt;https:<span class="hljs-comment">//login.skype.com/login/sso?go=webclient.xmm&amp;amp;docid=0-weu-d16-235cdecedb3e39f4ec1c2e451b3d211b&lt;/a&gt;&lt;/OriginalName&gt;&lt;/URIObject&gt;</span><br>&nbsp;</code></pre>';

    const m = msg({
      id: "code-block-1",
      from: "8:orgid:siphesihle-uuid",
      imdisplayname: "Siphesihle Thomo",
      messagetype: "RichText/Html",
      contenttype: "Text",
      content: realContent,
    });

    // Should be renderable (not filtered by isLikelySystemOrCallBlob)
    expect(isRenderableChatMessage(m)).toBe(true);

    const parts = messageRichPartsForDisplay(m);
    expect(parts).not.toBeNull();
    const bodyText = parts?.body.map((p) => p.text).join("") ?? "";
    expect(bodyText).toContain("It's just this");

    // Should produce a code_block part for the <pre><code> content
    const codeBlockParts =
      parts?.body.filter((p) => p.kind === "code_block") ?? [];
    expect(codeBlockParts.length).toBe(1);
    expect(codeBlockParts[0].language).toBe("json");
    expect(codeBlockParts[0].text).toContain("URIObject");
    expect(codeBlockParts[0].text).toContain("three-body.html");
  });

  it("parses inline <code> with code mark", () => {
    const m = msg({
      id: "inline-code-1",
      from: "8:orgid:test-uuid",
      imdisplayname: "Test User",
      messagetype: "RichText/Html",
      contenttype: "Text",
      content: "<p>Use <code>console.log()</code> to debug</p>",
    });
    const parts = messageRichPartsForDisplay(m);
    expect(parts).not.toBeNull();
    const codeParts =
      parts?.body.filter((p) => p.kind === "text" && p.code) ?? [];
    expect(codeParts.length).toBe(1);
    expect(codeParts[0].text).toBe("console.log()");
  });

  it("parses <pre> without language class", () => {
    const m = msg({
      id: "pre-no-lang",
      from: "8:orgid:test-uuid",
      imdisplayname: "Test User",
      messagetype: "RichText/Html",
      contenttype: "Text",
      content: "<pre><code>function hello() { return 42; }</code></pre>",
    });
    const parts = messageRichPartsForDisplay(m);
    expect(parts).not.toBeNull();
    const codeBlockParts =
      parts?.body.filter((p) => p.kind === "code_block") ?? [];
    expect(codeBlockParts.length).toBe(1);
    expect(codeBlockParts[0].text).toContain("function hello()");
    expect(codeBlockParts[0].language).toBeUndefined();
  });
});
