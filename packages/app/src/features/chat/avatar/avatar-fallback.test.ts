import { describe, expect, it } from "vitest";
import { avatarFallbackPresentation } from "./avatar-fallback";

describe("avatarFallbackPresentation", () => {
  it("uses name and surname initials", () => {
    expect(avatarFallbackPresentation("Daniel Makanda").initials).toBe("DM");
  });

  it("keeps the same gradient for the same initials", () => {
    expect(avatarFallbackPresentation("Daniel Makanda").style).toEqual(
      avatarFallbackPresentation("Diana Moss").style,
    );
  });

  it("varies the gradient by initials", () => {
    expect(avatarFallbackPresentation("Daniel Makanda").style).not.toEqual(
      avatarFallbackPresentation("Alex Smith").style,
    );
  });
});
