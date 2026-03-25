import { presenceBadgeClassName } from "@/lib/teams-presence";
import { cn } from "@/lib/utils";
import type { PresenceInfo } from "@/services/teams/types";

export function PresenceBadge({
  presence,
  size = "md",
  className,
}: {
  presence?: PresenceInfo | null;
  size?: "sm" | "md" | "lg";
  className?: string;
}) {
  const tone = presenceBadgeClassName(presence);
  const sizeClassName =
    size === "lg" ? "size-3.5" : size === "sm" ? "size-2.5" : "size-3";

  if (!tone) return null;

  return (
    <span
      className={cn(
        "absolute right-0 bottom-0 z-10 inline-flex shrink-0 rounded-full border-2 border-background shadow-sm",
        sizeClassName,
        tone,
        className,
      )}
      aria-hidden="true"
      data-slot="presence-badge"
    />
  );
}
