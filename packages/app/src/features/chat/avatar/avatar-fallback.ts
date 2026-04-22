import { initialsFromLabel } from "@better-teams/core/chat";
import type { CSSProperties } from "react";

type AvatarFallbackPresentation = {
  initials: string;
  style: CSSProperties;
};

function seedFromInitials(initials: string): number {
  return Array.from(initials).reduce(
    (seed, char) => seed * 37 + char.charCodeAt(0),
    17,
  );
}

export function avatarFallbackPresentation(
  label: string,
): AvatarFallbackPresentation {
  const initials = initialsFromLabel(label);
  const seed = seedFromInitials(initials);
  const primaryHue = (seed * 29) % 360;
  const secondaryHue = (primaryHue + 38 + (seed % 34)) % 360;
  const primarySaturation = 52 + (seed % 12);
  const secondarySaturation = 48 + ((seed >> 3) % 14);
  const primaryLightness = 32 + ((seed >> 2) % 8);
  const secondaryLightness = 43 + ((seed >> 5) % 10);

  return {
    initials,
    style: {
      backgroundColor: `hsl(${primaryHue} ${primarySaturation}% ${primaryLightness}%)`,
      backgroundImage: `linear-gradient(135deg, hsl(${primaryHue} ${primarySaturation}% ${primaryLightness}%) 0%, hsl(${secondaryHue} ${secondarySaturation}% ${secondaryLightness}%) 100%)`,
      color: "white",
      textShadow: "0 1px 1px rgb(0 0 0 / 0.18)",
    },
  };
}
