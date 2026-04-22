import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { render, screen } from "@testing-library/react";
import { describe, expect, it } from "vitest";
import { MessageRow, messageRowPropsEqual } from "../message/MessageRow";
import type { ProfileData } from "../profile/ProfileCard";
import type { DisplayMessage } from "../thread/types";

function makeEntry(): DisplayMessage {
  return {
    displayName: "Alex",
    self: false,
    time: "4:35 PM",
    message: {
      id: "m1",
      conversationId: "c1",
      type: "Message",
      messagetype: "Text",
      contenttype: "text",
      content: "Hello there",
      from: "8:orgid:alex",
      composetime: "2024-06-01T16:35:00.000Z",
      originalarrivaltime: "2024-06-01T16:35:00.000Z",
    },
    parts: {
      body: [{ kind: "text", text: "Hello there" }],
      quote: null,
      quoteRef: null,
      attachments: [],
    },
    deleted: false,
    edited: false,
    bodyPreview: "Hello there",
    searchText: "alex hello there",
  };
}

function renderMessageRow(element: React.ReactElement) {
  const queryClient = new QueryClient({
    defaultOptions: {
      queries: {
        retry: false,
      },
    },
  });

  return render(
    <QueryClientProvider client={queryClient}>{element}</QueryClientProvider>,
  );
}

describe("MessageRow", () => {
  it("keeps hover timestamps on one line", () => {
    renderMessageRow(<MessageRow entry={makeEntry()} showMeta={false} />);

    const timestamp = screen.getByText("4:35 PM");
    expect(timestamp.className).toContain("whitespace-nowrap");
  });

  it("renders sent and read details for self messages", () => {
    renderMessageRow(
      <MessageRow
        entry={{
          ...makeEntry(),
          self: true,
          displayName: "You",
          readStatus: "read",
          sentAt: "Jun 1, 4:35 PM",
          readAt: "Jun 1, 4:36 PM",
        }}
        showMeta
      />,
    );

    expect(screen.getByText("Sent")).toBeInTheDocument();
    expect(screen.getByText("Jun 1, 4:35 PM")).toBeInTheDocument();
    expect(screen.getByText("Read")).toBeInTheDocument();
    expect(screen.getByText("Jun 1, 4:36 PM")).toBeInTheDocument();
  });

  it("renders unread copy when the message is delivered but not read", () => {
    renderMessageRow(
      <MessageRow
        entry={{
          ...makeEntry(),
          self: true,
          displayName: "You",
          readStatus: "delivered",
          sentAt: "Jun 1, 4:35 PM",
          readAt: "",
        }}
        showMeta
      />,
    );

    expect(screen.getByText("Delivered, not read yet")).toBeInTheDocument();
  });

  it("renders aggregate group receipt details instead of a fake people list", () => {
    renderMessageRow(
      <MessageRow
        entry={{
          ...makeEntry(),
          self: true,
          displayName: "You",
          readStatus: "read",
          sentAt: "Jun 1, 4:35 PM",
          readAt: "Jun 1, 4:42 PM",
          receiptScope: "group",
          receiptSeenBy: [
            {
              mri: "8:orgid:sam",
              name: "Sam Jordan",
              readAt: "Jun 1, 4:42 PM",
            },
            {
              mri: "8:orgid:casey",
              name: "Casey Long",
              readAt: "Jun 1, 4:40 PM",
            },
          ],
          receiptUnseenBy: [
            {
              mri: "8:orgid:morgan",
              name: "Morgan Lee",
            },
          ],
        }}
        showMeta
      />,
    );

    expect(screen.getByText("Seen by")).toBeInTheDocument();
    expect(screen.getByText("Sam Jordan")).toBeInTheDocument();
    expect(screen.getByText("Casey Long")).toBeInTheDocument();
    expect(screen.getByText("Not seen by")).toBeInTheDocument();
    expect(screen.getByText("Morgan Lee")).toBeInTheDocument();
    expect(screen.getByText("Jun 1, 4:42 PM")).toBeInTheDocument();
  });

  it("skips rerender when parent recreates an equivalent profile object", () => {
    const entry = makeEntry();
    const baseProfile: ProfileData = {
      mri: "8:orgid:alex",
      displayName: "Alex",
      avatarThumbSrc: "https://example.com/thumb.png",
      avatarFullSrc: "https://example.com/full.png",
      email: "alex@example.com",
      jobTitle: "Engineer",
      department: "Product",
      companyName: "Acme",
      tenantName: "Acme",
      location: "Remote",
      currentConversationId: "c1",
      sharedConversationHeading: "Other chats with Alex",
      sharedConversations: [{ id: "c2", title: "Alex", kind: "dm" }],
      onOpenConversation: () => undefined,
      onMessage: () => undefined,
    };

    expect(
      messageRowPropsEqual(
        {
          entry,
          showMeta: true,
          profile: baseProfile,
        },
        {
          entry,
          showMeta: true,
          profile: {
            ...baseProfile,
            sharedConversations: [...(baseProfile.sharedConversations ?? [])],
            onOpenConversation: () => undefined,
            onMessage: () => undefined,
          },
        },
      ),
    ).toBe(true);
  });

  it("renders rich content immediately for stable row height", () => {
    renderMessageRow(
      <MessageRow
        entry={{
          ...makeEntry(),
          parts: {
            body: [{ kind: "text", text: "Loaded rich content" }],
            quote: null,
            quoteRef: null,
            attachments: [],
          },
          bodyPreview: "Preview copy",
        }}
        showMeta
      />,
    );

    expect(screen.getByText("Loaded rich content")).toBeInTheDocument();
  });

  it("keeps the avatar slot empty while fallback is not ready", () => {
    renderMessageRow(
      <MessageRow
        entry={{
          ...makeEntry(),
          displayName: "Daniel Makanda",
        }}
        showMeta
        avatarFallbackReady={false}
      />,
    );

    expect(screen.getByText("Daniel Makanda")).toBeInTheDocument();
    expect(screen.queryByText("DM")).not.toBeInTheDocument();
  });

  it("shows initials after fallback is ready", () => {
    renderMessageRow(
      <MessageRow
        entry={{
          ...makeEntry(),
          displayName: "Daniel Makanda",
        }}
        showMeta
        avatarFallbackReady
      />,
    );

    expect(screen.getByText("DM")).toBeInTheDocument();
  });

  it("renders attachment frames immediately for stable media rows", () => {
    renderMessageRow(
      <MessageRow
        entry={{
          ...makeEntry(),
          parts: {
            body: [{ kind: "text", text: "Loaded rich content" }],
            quote: null,
            quoteRef: null,
            attachments: [
              {
                kind: "image",
                title: "Screenshot.png",
                fileName: "Screenshot.png",
                objectUrl: "https://example.com/file",
                openUrl: "https://example.com/open",
                fileSize: 2048,
              },
            ],
          },
          bodyPreview: "Preview copy",
        }}
        showMeta
      />,
    );

    expect(screen.getByText("Screenshot.png")).toBeInTheDocument();
    expect(screen.getByText("2.0 KB")).toBeInTheDocument();
  });
});
