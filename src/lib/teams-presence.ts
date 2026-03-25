import type { PresenceInfo } from "@/services/teams/types";

function normalizePresenceToken(value: string | undefined): string {
  return (
    value
      ?.trim()
      .toLowerCase()
      .replace(/[^a-z]/g, "") ?? ""
  );
}

export function presenceLabel(presence?: PresenceInfo | null): string {
  if (!presence) return "Unknown";
  return presence.availability || presence.activity || "Unknown";
}

export function presenceDescription(presence?: PresenceInfo | null): string {
  if (!presence) return "Presence unknown";
  const status = presenceLabel(presence);
  const message = presence.statusMessage?.message?.trim();
  return message ? `${status}: ${message}` : status;
}

export function presenceBadgeClassName(
  presence?: PresenceInfo | null,
): string | null {
  const availability = normalizePresenceToken(presence?.availability);
  const activity = normalizePresenceToken(presence?.activity);
  const token = availability || activity;

  if (!token) return null;
  if (token.includes("available")) return "bg-emerald-500";
  if (token.includes("busy")) return "bg-rose-500";
  if (token.includes("donotdisturb") || token.includes("presenting")) {
    return "bg-rose-600";
  }
  if (
    token.includes("away") ||
    token.includes("berightback") ||
    token.includes("idle")
  ) {
    return "bg-amber-400";
  }
  if (token.includes("offline")) return "bg-zinc-400";
  return "bg-sky-500";
}
