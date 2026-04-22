import { describe, expect, it } from "vitest";
import { ConversationsResponseSchema, MessageSchema } from "./schemas";

describe("MessageSchema", () => {
  it("accepts numeric contenttype and timestamps from chat service", () => {
    const parsed = MessageSchema.parse({
      id: 1,
      conversationId: "19:thread@thread.v2",
      type: 0,
      messagetype: "Text",
      contenttype: 1,
      from: "8:orgid:abc",
      composetime: 1710800000000,
      originalarrivaltime: 1710800000001,
    });
    expect(parsed.contenttype).toBe("1");
    expect(parsed.id).toBe("1");
    expect(parsed.composetime).toBe("2024-03-18T22:13:20.000Z");
    expect(parsed.originalarrivaltime).toBe("2024-03-18T22:13:20.001Z");
    expect(parsed.timestamp).toBe("2024-03-18T22:13:20.001Z");
  });

  it("coerces null lastMessage string fields to empty strings", () => {
    const parsed = MessageSchema.parse({
      id: null,
      conversationid: "19:x@thread.v2",
      type: null,
      messagetype: null,
      contenttype: null,
      from: null,
      composetime: null,
      originalarrivaltime: null,
    });
    expect(parsed.id).toBe("");
    expect(parsed.type).toBe("");
    expect(parsed.contenttype).toBe("");
    expect(parsed.timestamp).toBe("");
  });

  it("normalizes sender and known json properties", () => {
    const parsed = MessageSchema.parse({
      id: "m1",
      conversationid: "19:x@thread.v2",
      type: "Message",
      messagetype: "Text",
      contenttype: "Text",
      from: "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/contacts/8:orgid:abc-123",
      imdisplayname: "  Pat  ",
      composetime: "2024-01-01T00:00:00Z",
      originalarrivaltime: "2024-01-01T00:00:00Z",
      properties: {
        qtdMsgs: '[{"messageId":1,"sender":"8:orgid:peer"}]',
        emotions: '[{"key":"like","users":[{"mri":"8:orgid:peer"}]}]',
      },
    });
    expect(parsed.fromMri).toBe("8:orgid:abc-123");
    expect(parsed.senderDisplayName).toBe("Pat");
    expect(parsed.properties?.qtdMsgs).toEqual([
      { messageId: 1, sender: "8:orgid:peer" },
    ]);
  });
});

describe("ConversationsResponseSchema", () => {
  it("parses conversations with loosely typed lastMessage", () => {
    const res = ConversationsResponseSchema.parse({
      conversations: [
        {
          id: "19:a@thread.v2",
          lastMessage: {
            id: "m1",
            type: "Message",
            messagetype: "Text",
            contenttype: 2,
            from: "8:orgid:x",
            composetime: "2024-01-01T00:00:00Z",
            originalarrivaltime: "2024-01-01T00:00:00Z",
            conversationId: "19:a@thread.v2",
          },
        },
      ],
    });
    expect(res.conversations[0]?.lastMessage?.contenttype).toBe("2");
  });

  it("uses parent conversation id when lastMessage omits conversationId", () => {
    const res = ConversationsResponseSchema.parse({
      conversations: [
        {
          id: "19:parent@thread.v2",
          lastMessage: {
            id: "m1",
            type: "Message",
            messagetype: "Text",
            contenttype: "text",
            from: "8:orgid:x",
            composetime: "2024-01-01T00:00:00Z",
            originalarrivaltime: "2024-01-01T00:00:00Z",
          },
        },
      ],
    });
    expect(res.conversations[0]?.lastMessage?.conversationId).toBe(
      "19:parent@thread.v2",
    );
  });

  it("derives memberCount and lastActivityTime from the wire payload", () => {
    const res = ConversationsResponseSchema.parse({
      conversations: [
        {
          id: "19:parent@thread.v2",
          threadProperties: { membercount: "2", memberCount: 5 },
          members: [
            { id: "8:orgid:a" },
            { id: "8:orgid:b" },
            { id: "8:orgid:c" },
          ],
          properties: { consumptionhorizon: "1;2;3" },
          lastMessage: {
            id: "m1",
            type: "Message",
            messagetype: "Text",
            contenttype: "text",
            from: "8:orgid:x",
            composetime: 1710800000000,
            originalarrivaltime: 1710800001000,
          },
        },
      ],
    });
    expect(res.conversations[0]?.memberCount).toBe(5);
    expect(res.conversations[0]?.consumptionHorizon).toBe("1;2;3");
    expect(res.conversations[0]?.lastActivityTime).toBe(
      "2024-03-18T22:13:21.000Z",
    );
  });

  it("normalizes known conversation property blobs", () => {
    const res = ConversationsResponseSchema.parse({
      conversations: [
        {
          id: "19:parent@thread.v2",
          properties: {
            meetingInfo: '{"meetingId":"abc"}',
            alerts: '[{"type":"mention"}]',
          },
          lastMessage: {
            id: "m1",
            type: "Message",
            messagetype: "Text",
            contenttype: "text",
            from: "8:orgid:x",
            composetime: "2024-01-01T00:00:00Z",
            originalarrivaltime: "2024-01-01T00:00:00Z",
          },
        },
      ],
    });
    expect(res.conversations[0]?.properties?.meetingInfo).toEqual({
      meetingId: "abc",
    });
    expect(res.conversations[0]?.properties?.alerts).toEqual([
      { type: "mention" },
    ]);
  });

  it("tags known special thread ids", () => {
    const res = ConversationsResponseSchema.parse({
      conversations: [
        {
          id: "48:notifications",
          lastMessage: {
            id: "m1",
            type: "Message",
            messagetype: "Text",
            contenttype: "text",
            from: "8:orgid:x",
            composetime: "2024-01-01T00:00:00Z",
            originalarrivaltime: "2024-01-01T00:00:00Z",
          },
        },
      ],
    });
    expect(res.conversations[0]?.specialThreadType).toBe("notifications");
  });
});
