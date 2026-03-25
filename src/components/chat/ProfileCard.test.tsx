import { fireEvent, render, screen } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import { type ProfileData, ProfilePanel, ProfileTrigger } from "./ProfileCard";

function makeProfile(overrides: Partial<ProfileData> = {}): ProfileData {
  return {
    mri: "8:orgid:test",
    displayName: "Alex Johnson",
    currentConversationId: "c1",
    ...overrides,
  };
}

describe("ProfilePanel", () => {
  it("uses the message action and removes the dead more controls", () => {
    const onMessage = vi.fn();

    render(
      <ProfilePanel
        profile={makeProfile({ onMessage })}
        onClose={() => undefined}
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: /^message$/i }));

    expect(onMessage).toHaveBeenCalledTimes(1);
    expect(
      screen.queryByRole("button", { name: /more options/i }),
    ).not.toBeInTheDocument();
    expect(
      screen.queryByRole("button", { name: /^more$/i }),
    ).not.toBeInTheDocument();
  });

  it("hides the message action when no alternate chat is available", () => {
    render(
      <ProfilePanel
        profile={makeProfile({ onMessage: undefined })}
        onClose={() => undefined}
      />,
    );

    expect(
      screen.queryByRole("button", { name: /^message$/i }),
    ).not.toBeInTheDocument();
  });

  it("does not show a misleading shared chat count when the list is truncated", () => {
    render(
      <ProfilePanel
        profile={makeProfile({
          sharedConversations: Array.from({ length: 8 }, (_, index) => ({
            id: `c${index + 2}`,
            title: `Chat ${index + 1}`,
            kind: "group",
          })),
        })}
        onClose={() => undefined}
      />,
    );

    expect(screen.getByText("OTHER CHATS")).toBeInTheDocument();
    expect(screen.queryByText(/^8$/)).not.toBeInTheDocument();
  });
});

describe("ProfileTrigger", () => {
  it("portals the profile drawer to the document body", () => {
    render(
      <div data-testid="host">
        <ProfileTrigger profile={makeProfile()}>
          <span>Open profile</span>
        </ProfileTrigger>
      </div>,
    );

    fireEvent.click(screen.getByRole("button", { name: /view profile/i }));

    const dialog = screen.getByRole("dialog", {
      name: "Profile: Alex Johnson",
    });

    expect(dialog.parentElement).toBe(document.body);
    expect(screen.getByTestId("host")).not.toContainElement(dialog);
  });
});
