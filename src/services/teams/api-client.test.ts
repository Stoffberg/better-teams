import { describe, expect, it, vi } from "vitest";
import { cacheImageFile, filePathToAssetUrl } from "@/lib/electron-bridge";
import { TeamsApiClient } from "./api-client";

vi.mock("@/lib/electron-bridge", () => ({
  cacheImageFile: vi.fn(),
  extractTokens: vi.fn().mockResolvedValue([]),
  filePathToAssetUrl: vi.fn((filePath: string) => `asset://${filePath}`),
  getAuthToken: vi.fn().mockResolvedValue(null),
}));

describe("TeamsApiClient.getPresence", () => {
  it("maps nested presence responses back to the requested MRIs", async () => {
    const fetchImpl = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify([
          {
            mri: "8:orgid:peer-1",
            presence: {
              availability: "Available",
              activity: "Available",
            },
            status: 20000,
          },
          {
            mri: "8:orgid:peer-2",
            presence: {
              availability: "Offline",
              activity: "Offline",
            },
            status: 20000,
          },
        ]),
        { status: 200 },
      ),
    );

    const client = new TeamsApiClient(undefined, { fetchImpl });
    Reflect.set(client, "authToken", {
      token: "bearer-token",
      expiresAt: new Date("2999-01-01T00:00:00Z"),
    });
    Reflect.set(client, "skypeToken", "skype-token");
    Reflect.set(client, "regionGtms", {
      unifiedPresence: "https://presence.teams.microsoft.com",
    });

    const presence = await client.getPresence([
      "8:orgid:peer-1",
      "8:orgid:peer-2",
    ]);

    expect(fetchImpl).toHaveBeenCalledWith(
      "https://presence.teams.microsoft.com/v1/presence/getpresence",
      expect.objectContaining({
        method: "POST",
        body: JSON.stringify([
          { mri: "8:orgid:peer-1" },
          { mri: "8:orgid:peer-2" },
        ]),
      }),
    );
    expect(presence).toEqual({
      "8:orgid:peer-1": {
        availability: "Available",
        activity: "Available",
      },
      "8:orgid:peer-2": {
        availability: "Offline",
        activity: "Offline",
      },
    });
  });

  it("stops retrying presence after the first unauthorized response", async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(new Response(null, { status: 401 }));

    const client = new TeamsApiClient(undefined, { fetchImpl });
    Reflect.set(client, "authToken", {
      token: "bearer-token",
      expiresAt: new Date("2999-01-01T00:00:00Z"),
    });
    Reflect.set(client, "skypeToken", "skype-token");
    Reflect.set(client, "regionGtms", {
      unifiedPresence: "https://presence.teams.microsoft.com",
    });

    await expect(client.getPresence(["8:orgid:peer-1"])).resolves.toEqual({});
    await expect(client.getPresence(["8:orgid:peer-2"])).resolves.toEqual({});

    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });
});

describe("TeamsApiClient.getAllConversations", () => {
  it("follows backwardLink conversation pagination URLs", async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            conversations: [
              {
                id: "19:first@thread.v2",
                lastMessage: {
                  id: "m1",
                  type: "Message",
                  messagetype: "Text",
                  contenttype: "text",
                  from: "8:orgid:first",
                  composetime: "2024-01-02T00:00:00Z",
                  originalarrivaltime: "2024-01-02T00:00:00Z",
                },
              },
            ],
            _metadata: {
              backwardLink:
                "https://emea.ng.msg.teams.microsoft.com/v1/users/8:orgid:test/conversations?view=msnp24Equivalent&startTime=123",
            },
          }),
          { status: 200 },
        ),
      )
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            conversations: [
              {
                id: "19:second@thread.v2",
                lastMessage: {
                  id: "m2",
                  type: "Message",
                  messagetype: "Text",
                  contenttype: "text",
                  from: "8:orgid:second",
                  composetime: "2024-01-01T00:00:00Z",
                  originalarrivaltime: "2024-01-01T00:00:00Z",
                },
              },
            ],
            _metadata: {},
          }),
          { status: 200 },
        ),
      );

    const client = new TeamsApiClient(undefined, { fetchImpl });
    Reflect.set(client, "skypeToken", "test-token");
    Reflect.set(client, "authToken", {
      expiresAt: new Date("2999-01-01T00:00:00Z"),
    });
    Reflect.set(client, "regionGtms", {
      chatService: "https://emea.ng.msg.teams.microsoft.com",
    });

    const response = await client.getAllConversations(100);

    expect(fetchImpl).toHaveBeenNthCalledWith(
      1,
      "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations?view=msnp24Equivalent&pageSize=100&startTime=0",
      {
        headers: {
          Authentication: "skypetoken=test-token",
        },
      },
    );
    expect(fetchImpl).toHaveBeenNthCalledWith(
      2,
      "https://emea.ng.msg.teams.microsoft.com/v1/users/8:orgid:test/conversations?view=msnp24Equivalent&startTime=123",
      {
        headers: {
          Authentication: "skypetoken=test-token",
        },
      },
    );
    expect(
      response.conversations.map((conversation) => conversation.id),
    ).toEqual(["19:first@thread.v2", "19:second@thread.v2"]);
  });
});

describe("TeamsApiClient.getFavoriteConversations", () => {
  it("requests the Teams favorites view", async () => {
    const fetchImpl = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          conversations: [
            {
              id: "19:favorite@thread.v2",
              properties: {
                favorite: true,
              },
              lastMessage: {
                id: "m1",
                type: "Message",
                messagetype: "Text",
                contenttype: "text",
                from: "8:orgid:first",
                composetime: "2024-01-02T00:00:00Z",
                originalarrivaltime: "2024-01-02T00:00:00Z",
              },
            },
          ],
          _metadata: {},
        }),
        { status: 200 },
      ),
    );

    const client = new TeamsApiClient(undefined, { fetchImpl });
    Reflect.set(client, "skypeToken", "test-token");
    Reflect.set(client, "authToken", {
      expiresAt: new Date("2999-01-01T00:00:00Z"),
    });
    Reflect.set(client, "regionGtms", {
      chatService: "https://emea.ng.msg.teams.microsoft.com",
    });

    const response = await client.getFavoriteConversations(50);

    expect(fetchImpl).toHaveBeenCalledWith(
      "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations?view=favorites&pageSize=50&startTime=0",
      {
        headers: {
          Authentication: "skypetoken=test-token",
        },
      },
    );
    expect(
      response.conversations.map((conversation) => conversation.id),
    ).toEqual(["19:favorite@thread.v2"]);
  });
});

describe("TeamsApiClient.setConversationFavorite", () => {
  it("updates the favorite property through the conversation properties endpoint", async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValue(new Response(null, { status: 200 }));

    const client = new TeamsApiClient(undefined, { fetchImpl });
    Reflect.set(client, "skypeToken", "test-token");
    Reflect.set(client, "authToken", {
      expiresAt: new Date("2999-01-01T00:00:00Z"),
    });
    Reflect.set(client, "regionGtms", {
      chatService: "https://emea.ng.msg.teams.microsoft.com",
    });

    await client.setConversationFavorite("19:test@thread.v2", true);

    expect(fetchImpl).toHaveBeenCalledWith(
      "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19%3Atest%40thread.v2/properties",
      {
        method: "PUT",
        headers: {
          Authentication: "skypetoken=test-token",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ favorite: true }),
      },
    );
  });
});

describe("TeamsApiClient.sendAttachmentMessage", () => {
  it("creates an AMS object, uploads the file, and posts a file message", async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ id: "0-123" }), { status: 200 }),
      )
      .mockResolvedValueOnce(new Response(null, { status: 200 }))
      .mockResolvedValueOnce(new Response(null, { status: 201 }));

    const client = new TeamsApiClient(undefined, { fetchImpl });
    Reflect.set(client, "skypeToken", "test-token");
    Reflect.set(client, "authToken", {
      expiresAt: new Date("2999-01-01T00:00:00Z"),
    });
    Reflect.set(client, "regionGtms", {
      ams: "https://api.asm.skype.com",
      chatService: "https://emea.ng.msg.teams.microsoft.com",
    });

    const file = new File(["hello"], "brief.txt", { type: "text/plain" });
    await client.sendAttachmentMessage("19:test@thread.v2", file, "Me", [
      "8:orgid:peer",
    ]);

    expect(fetchImpl).toHaveBeenNthCalledWith(
      1,
      "https://api.asm.skype.com/v1/objects",
      {
        method: "POST",
        headers: {
          Authorization: "skype_token test-token",
          "Content-Type": "application/json",
          "X-Client-Version": "0/0.0.0.0",
        },
        body: JSON.stringify({
          type: "sharing/file",
          permissions: {
            "8:orgid:peer": ["read"],
          },
          filename: "brief.txt",
        }),
      },
    );
    expect(fetchImpl).toHaveBeenNthCalledWith(
      2,
      "https://api.asm.skype.com/v1/objects/0-123/content/original",
      {
        method: "PUT",
        headers: {
          Authorization: "skype_token test-token",
          "Content-Type": "text/plain",
        },
        body: expect.any(ArrayBuffer),
      },
    );
    expect(fetchImpl).toHaveBeenNthCalledWith(
      3,
      "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19%3Atest%40thread.v2/messages",
      expect.objectContaining({
        method: "POST",
        headers: {
          Authentication: "skypetoken=test-token",
          "Content-Type": "application/json",
        },
      }),
    );
    const [, , messageCall] = fetchImpl.mock.calls;
    const payload = JSON.parse(messageCall[1].body);
    expect(payload.messagetype).toBe("RichText/Media_GenericFile");
    expect(payload.contenttype).toBe("RichText/Media_GenericFile");
    expect(payload.amsreferences).toEqual(["0-123"]);
    expect(payload.properties).toEqual({
      formatVariant: "RichText/Media_GenericFile",
    });
    expect(payload.content).toContain('type="File.1"');
    expect(payload.content).toContain("brief.txt");
  });
});

describe("TeamsApiClient profile image caching", () => {
  it("returns cached image paths as asset urls without refetching", async () => {
    const fetchImpl = vi.fn();
    const client = new TeamsApiClient(undefined, {
      fetchImpl,
      getCachedImagePath: vi.fn().mockResolvedValue("/tmp/avatar.png"),
    });
    Reflect.set(client, "authToken", {
      token: "bearer-token",
      expiresAt: new Date("2999-01-01T00:00:00Z"),
    });
    Reflect.set(client, "skypeToken", "skype-token");
    Reflect.set(client, "regionGtms", {
      middleTier: "https://teams.microsoft.com/api/mt/part/amer-03",
    });

    const imageSrc = await Reflect.get(
      client,
      "fetchAuthenticatedImageSrc",
    ).call(client, "https://cdn.example.com/avatar.png");

    expect(imageSrc).toBe("asset:///tmp/avatar.png");
    expect(filePathToAssetUrl).toHaveBeenCalledWith("/tmp/avatar.png");
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it("writes fetched avatars to disk and stores the file path", async () => {
    vi.mocked(cacheImageFile).mockResolvedValueOnce("/tmp/avatar.jpg");
    const setCachedImagePath = vi.fn().mockResolvedValue(undefined);
    const fetchImpl = vi.fn().mockResolvedValue(
      new Response(Uint8Array.from([0xff, 0xd8, 0xff, 0xdb]), {
        status: 200,
        headers: {
          "content-type": "image/jpeg",
        },
      }),
    );
    const client = new TeamsApiClient(undefined, {
      fetchImpl,
      getCachedImagePath: vi.fn().mockResolvedValue(null),
      setCachedImagePath,
    });
    Reflect.set(client, "authToken", {
      token: "bearer-token",
      expiresAt: new Date("2999-01-01T00:00:00Z"),
    });
    Reflect.set(client, "skypeToken", "skype-token");
    Reflect.set(client, "regionGtms", {
      middleTier: "https://teams.microsoft.com/api/mt/part/amer-03",
    });

    const imageSrc = await Reflect.get(
      client,
      "fetchAuthenticatedImageSrc",
    ).call(client, "https://cdn.example.com/avatar.jpg");

    expect(cacheImageFile).toHaveBeenCalledWith(
      "https://cdn.example.com/avatar.jpg",
      expect.any(Uint8Array),
      "jpg",
    );
    expect(setCachedImagePath).toHaveBeenCalledWith(
      "https://cdn.example.com/avatar.jpg",
      "/tmp/avatar.jpg",
    );
    expect(imageSrc).toBe("asset:///tmp/avatar.jpg");
  });

  it("uses skype_token authorization for asm image views", async () => {
    vi.mocked(cacheImageFile).mockResolvedValueOnce("/tmp/preview.jpg");
    const fetchImpl = vi.fn().mockResolvedValue(
      new Response(Uint8Array.from([0xff, 0xd8, 0xff, 0xdb]), {
        status: 200,
        headers: {
          "content-type": "image/jpeg",
        },
      }),
    );
    const client = new TeamsApiClient(undefined, {
      fetchImpl,
      getCachedImagePath: vi.fn().mockResolvedValue(null),
      setCachedImagePath: vi.fn().mockResolvedValue(undefined),
    });
    Reflect.set(client, "authToken", {
      token: "bearer-token",
      expiresAt: new Date("2999-01-01T00:00:00Z"),
    });
    Reflect.set(client, "skypeToken", "skype-token");
    Reflect.set(client, "regionGtms", {
      middleTier: "https://teams.microsoft.com/api/mt/part/amer-03",
    });

    await Reflect.get(client, "fetchAuthenticatedImageSrc").call(
      client,
      "https://eu-api.asm.skype.com/v1/objects/0-123/views/imgo",
    );

    expect(fetchImpl).toHaveBeenCalledWith(
      "https://eu-api.asm.skype.com/v1/objects/0-123/views/imgo",
      {
        headers: {
          Authorization: "skype_token skype-token",
          Accept:
            "image/avif,image/webp,image/apng,image/png,image/*,*/*;q=0.8",
        },
        redirect: "follow",
      },
    );
  });
});
