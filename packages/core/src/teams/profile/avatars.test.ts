import { describe, expect, it } from "vitest";
import type { Conversation, Message } from "../types";
import {
  applyProfileDisplayNameToRowMrIs,
  applyProfilePhotoDataUrlToRowMrIs,
  canonAvatarMri,
  collectProfileAvatarMris,
  collectSkypeMriLikeStringsFromJson,
  displayNameFromShortProfileRow,
  dmConversationAvatarMri,
  normalizeFetchShortProfileRows,
  shortProfileRowToMriAndImageUrl,
} from "./avatars";

describe("dmConversationAvatarMri", () => {
  it("prefers a non-self member MRI in a DM", () => {
    const c: Conversation = {
      id: "x",
      threadProperties: { membercount: "2", threadType: "single" },
      members: [
        { id: "8:self", role: "User", isMri: true },
        { id: "8:orgid:aaa", role: "User", isMri: true },
      ],
    };
    expect(dmConversationAvatarMri(c, "self")).toBe("8:orgid:aaa");
  });

  it("falls back to lastMessage.from when not self", () => {
    const c: Conversation = {
      id: "x",
      threadProperties: { membercount: "2" },
      lastMessage: {
        id: "m",
        from: "8:peer",
        conversationId: "x",
        type: "Message",
        messagetype: "Text",
        contenttype: "text",
        composetime: "",
        originalarrivaltime: "",
      },
    };
    expect(dmConversationAvatarMri(c, "me")).toBe("8:peer");
  });

  it("extracts MRI from contact URL forms", () => {
    const c: Conversation = {
      id: "x",
      threadProperties: { membercount: "2" },
      lastMessage: {
        id: "m",
        from: "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/contacts/8:orgid:peer-id",
        conversationId: "x",
        type: "Message",
        messagetype: "Text",
        contenttype: "text",
        composetime: "",
        originalarrivaltime: "",
      },
    };
    expect(dmConversationAvatarMri(c, "me")).toBe("8:orgid:peer-id");
  });

  it("extracts peer MRI from dm conversation id pair", () => {
    const c: Conversation = {
      id: "19:48bb84f7-e080-4d88-8043-ae0e8cf745fc_f4cc62d6-05d5-48b0-9feb-ffe47197d860@unq.gbl.spaces",
      threadProperties: { membercount: "2" },
    };
    expect(
      dmConversationAvatarMri(
        c,
        "8:orgid:48bb84f7-e080-4d88-8043-ae0e8cf745fc",
      ),
    ).toBe("8:orgid:f4cc62d6-05d5-48b0-9feb-ffe47197d860");
  });
});

describe("collectProfileAvatarMris", () => {
  it("dedupes and includes self MRI when skypeId is set", () => {
    const m: Message = {
      id: "1",
      from: "8:peer",
      conversationId: "c",
      type: "Message",
      messagetype: "Text",
      contenttype: "text",
      composetime: "",
      originalarrivaltime: "",
    };
    const list = collectProfileAvatarMris({
      conversations: [],
      messages: [m, { ...m, id: "2" }],
      selfSkypeId: "self",
    });
    expect(list).toEqual(["8:self", "8:peer"]);
  });

  it("extracts MRI from URL sender values in messages", () => {
    const m: Message = {
      id: "1",
      from: "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/contacts/8:orgid:peer",
      conversationId: "c",
      type: "Message",
      messagetype: "Text",
      contenttype: "text",
      composetime: "",
      originalarrivaltime: "",
    };
    const list = collectProfileAvatarMris({
      conversations: [],
      messages: [m],
      selfSkypeId: "self",
    });
    expect(list).toEqual(["8:self", "8:orgid:peer"]);
  });

  it("normalizes selfSkypeId that already includes 8 prefix", () => {
    const list = collectProfileAvatarMris({
      conversations: [],
      messages: [],
      selfSkypeId: "8:orgid:self-id",
    });
    expect(list).toEqual(["8:orgid:self-id"]);
  });

  it("adds participant MRIs parsed from dm conversation ids", () => {
    const list = collectProfileAvatarMris({
      conversations: [
        {
          id: "19:48bb84f7-e080-4d88-8043-ae0e8cf745fc_f4cc62d6-05d5-48b0-9feb-ffe47197d860@unq.gbl.spaces",
        } as Conversation,
      ],
      messages: [],
      selfSkypeId: "8:orgid:48bb84f7-e080-4d88-8043-ae0e8cf745fc",
    });
    expect(list).toContain("8:orgid:f4cc62d6-05d5-48b0-9feb-ffe47197d860");
  });

  it("adds dm member ids as avatar candidates", () => {
    const list = collectProfileAvatarMris({
      conversations: [
        {
          id: "19:dm@unq.gbl.spaces",
          threadProperties: { membercount: "2", threadType: "single" },
          members: [
            { id: "8:orgid:self-id", role: "User", isMri: true },
            { id: "8:orgid:peer-id", role: "User", isMri: true },
          ],
        } as Conversation,
      ],
      messages: [],
      selfSkypeId: "8:orgid:self-id",
    });
    expect(list).toContain("8:orgid:peer-id");
  });

  it("adds group member ids as avatar candidates", () => {
    const list = collectProfileAvatarMris({
      conversations: [
        {
          id: "19:group@thread.v2",
          threadProperties: { membercount: "3", threadType: "chat" },
          members: [
            { id: "8:orgid:self-id", role: "User", isMri: true },
            { id: "8:orgid:peer-a", role: "User", isMri: true },
            { id: "8:orgid:peer-b", role: "User", isMri: true },
          ],
        } as Conversation,
      ],
      messages: [],
      selfSkypeId: "8:orgid:self-id",
    });

    expect(list).toEqual([
      "8:orgid:self-id",
      "8:orgid:peer-a",
      "8:orgid:peer-b",
    ]);
  });
});

describe("normalizeFetchShortProfileRows", () => {
  it("unwraps common envelope shapes", () => {
    expect(normalizeFetchShortProfileRows([1])).toEqual([1]);
    expect(normalizeFetchShortProfileRows({ value: [2] })).toEqual([2]);
    expect(normalizeFetchShortProfileRows({ profiles: [3] })).toEqual([3]);
  });

  it("flattens MRI-keyed profile maps", () => {
    const rows = normalizeFetchShortProfileRows({
      "8:orgid:aaa": { displayName: "A", imageUri: "https://x/y.jpg" },
    });
    expect(rows).toHaveLength(1);
    expect(shortProfileRowToMriAndImageUrl(rows[0])).toEqual({
      mri: "8:orgid:aaa",
      imageUrl: "https://x/y.jpg",
    });
  });
});

describe("collectSkypeMriLikeStringsFromJson", () => {
  it("finds nested orgid and live MRIs", () => {
    const found = collectSkypeMriLikeStringsFromJson({
      skypeTeamsInfo: {
        userMri: "8:live:.cid.abc",
        linked: ["8:orgid:1111-2222"],
      },
    });
    expect(found).toEqual(
      expect.arrayContaining(["8:live:.cid.abc", "8:orgid:1111-2222"]),
    );
    expect(found.length).toBe(2);
  });

  it("extracts MRI from contact URLs", () => {
    const found = collectSkypeMriLikeStringsFromJson({
      sender:
        "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/contacts/8:orgid:abc-123",
    });
    expect(found).toEqual(["8:orgid:abc-123"]);
  });
});

describe("displayNameFromShortProfileRow", () => {
  it("reads displayName", () => {
    expect(
      displayNameFromShortProfileRow({
        mri: "8:orgid:x",
        displayName: "Alex",
      }),
    ).toBe("Alex");
  });

  it("joins givenName and surname", () => {
    expect(
      displayNameFromShortProfileRow({
        mri: "8:orgid:x",
        givenName: "Lee",
        surname: "Kim",
      }),
    ).toBe("Lee Kim");
  });
});

describe("applyProfileDisplayNameToRowMrIs", () => {
  it("stores under every MRI found in the profile row", () => {
    const map: Record<string, string> = {};
    const setFor = (mri: string, d: string) => {
      map[canonAvatarMri(mri)] = d;
    };
    applyProfileDisplayNameToRowMrIs(
      {
        mri: "8:orgid:aaa",
        nested: { alt: "8:live:.cid.x" },
        displayName: "Pat",
      },
      "Pat",
      setFor,
    );
    expect(map[canonAvatarMri("8:orgid:aaa")]).toBe("Pat");
    expect(map[canonAvatarMri("8:live:.cid.x")]).toBe("Pat");
  });
});

describe("applyProfilePhotoDataUrlToRowMrIs", () => {
  it("stores under every MRI found in the profile row", () => {
    const map: Record<string, string> = {};
    const setFor = (mri: string, d: string) => {
      map[mri] = d;
    };
    applyProfilePhotoDataUrlToRowMrIs(
      {
        mri: "8:orgid:aaa",
        nested: { alt: "8:live:.cid.x" },
        imageUri: "https://x/p.jpg",
      },
      "data:image/png;base64,QQ==",
      setFor,
    );
    expect(map["8:orgid:aaa"]).toBe("data:image/png;base64,QQ==");
    expect(map["8:live:.cid.x"]).toBe("data:image/png;base64,QQ==");
  });
});

describe("shortProfileRowToMriAndImageUrl", () => {
  it("reads imageUri and mri", () => {
    expect(
      shortProfileRowToMriAndImageUrl({
        mri: "8:orgid:x",
        imageUri: "https://example.com/p.jpg",
      }),
    ).toEqual({
      mri: "8:orgid:x",
      imageUrl: "https://example.com/p.jpg",
    });
  });

  it("accepts relative image paths", () => {
    expect(
      shortProfileRowToMriAndImageUrl({
        mri: "8:live:y",
        imageUri: "/api/me/photo",
      }),
    ).toEqual({ mri: "8:live:y", imageUrl: "/api/me/photo" });
  });
});

describe("canonAvatarMri", () => {
  it("canonicalizes contact URL MRI keys", () => {
    expect(
      canonAvatarMri(
        "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/contacts/8:ORGID:ABC-123",
      ),
    ).toBe("8:orgid:abc-123");
  });
});
