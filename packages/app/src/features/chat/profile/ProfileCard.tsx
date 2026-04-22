import { presenceDescription } from "@better-teams/core/teams/presence";
import type { PresenceInfo } from "@better-teams/core/teams/types";
import { cn } from "@better-teams/ui/utils";
import { Hash, MessageSquare, Video, X } from "lucide-react";
import {
  type KeyboardEventHandler,
  type MouseEventHandler,
  type ReactNode,
  useState,
} from "react";
import { avatarFallbackPresentation } from "../avatar/avatar-fallback";
import { PresenceBadge } from "../presence/PresenceBadge";

export type ProfileData = {
  mri: string;
  displayName: string;
  avatarThumbSrc?: string;
  avatarFullSrc?: string;
  avatarFallbackReady?: boolean;
  email?: string;
  jobTitle?: string;
  department?: string;
  companyName?: string;
  tenantName?: string;
  location?: string;
  presence?: PresenceInfo | null;
  currentConversationId?: string;
  sharedConversationHeading?: string;
  sharedConversationsLoading?: boolean;
  onOpenConversation?: (conversationId: string) => void;
  onMessage?: () => void;
  sharedConversations?: Array<{
    id: string;
    title: string;
    kind: "dm" | "group" | "meeting";
    preview?: string;
    sideTime?: string;
  }>;
};

function profileOrgName(profile: ProfileData): string | undefined {
  return profile.tenantName || profile.companyName;
}

function ProfileAvatar({
  src,
  name,
  size,
  presence,
  fallbackReady = true,
}: {
  src?: string;
  name: string;
  size: "sm" | "lg" | "xl";
  presence?: PresenceInfo | null;
  fallbackReady?: boolean;
}) {
  const [failed, setFailed] = useState(false);
  const fallback = avatarFallbackPresentation(name);
  const cls = size === "xl" ? "size-28" : size === "lg" ? "size-24" : "size-10";
  const textCls =
    size === "xl"
      ? "text-3xl font-bold"
      : size === "lg"
        ? "text-2xl font-bold"
        : "text-[12px] font-semibold";
  const radius =
    size === "xl" || size === "lg" ? "rounded-full" : "rounded-2xl";

  if (src && !failed) {
    return (
      <span className="relative block">
        <img
          src={src}
          alt={name}
          onError={() => setFailed(true)}
          className={cn(cls, radius, "object-cover shadow-md")}
        />
        <PresenceBadge presence={presence} size="lg" />
      </span>
    );
  }
  if (!fallbackReady && !failed) {
    return (
      <span className="relative block" aria-hidden>
        <span className={cn("block shrink-0", cls, radius)} />
      </span>
    );
  }
  return (
    <span className="relative block">
      <span
        className={cn(
          "flex shrink-0 items-center justify-center shadow-md",
          cls,
          radius,
          textCls,
        )}
        style={fallback.style}
      >
        {fallback.initials}
      </span>
      <PresenceBadge presence={presence} size="lg" />
    </span>
  );
}

function HoverProfileCard({ profile }: { profile: ProfileData }) {
  const orgName = profileOrgName(profile);
  return (
    <div className="flex items-center gap-3 p-1">
      <ProfileAvatar
        src={profile.avatarFullSrc ?? profile.avatarThumbSrc}
        name={profile.displayName}
        size="sm"
        presence={profile.presence}
        fallbackReady={profile.avatarFallbackReady}
      />
      <div className="min-w-0">
        <p className="truncate text-[13px] font-semibold leading-tight text-foreground">
          {profile.displayName}
        </p>
        {profile.jobTitle ? (
          <p className="truncate pt-0.5 text-[11px] leading-tight text-muted-foreground">
            {profile.jobTitle}
          </p>
        ) : null}
        {profile.presence ? (
          <p className="truncate pt-0.5 text-[11px] leading-tight text-muted-foreground/80">
            {presenceDescription(profile.presence)}
          </p>
        ) : null}
        {orgName ? (
          <p className="truncate pt-0.5 text-[11px] leading-tight text-muted-foreground/80">
            {orgName}
          </p>
        ) : null}
      </div>
    </div>
  );
}

function SharedConversationIcon({
  kind,
}: {
  kind: "dm" | "group" | "meeting";
}) {
  if (kind === "meeting")
    return <Video className="size-3.5 shrink-0 text-muted-foreground" />;
  if (kind === "group")
    return <Hash className="size-3.5 shrink-0 text-muted-foreground" />;
  return <MessageSquare className="size-3.5 shrink-0 text-muted-foreground" />;
}

function SharedConversationList({ profile }: { profile: ProfileData }) {
  const items = (profile.sharedConversations ?? []).filter(
    (item) => item.id !== profile.currentConversationId,
  );
  if (!profile.sharedConversationsLoading && items.length === 0) return null;

  return (
    <div className="px-6 pt-4" aria-busy={profile.sharedConversationsLoading}>
      <div className="mb-1">
        <h5 className="text-[11px] font-semibold tracking-wide text-muted-foreground uppercase">
          OTHER CHATS
        </h5>
      </div>
      <div className="max-h-56 space-y-0.5 overflow-y-auto pr-1">
        {profile.sharedConversationsLoading
          ? Array.from({ length: 4 }, (_, index) => (
              <div
                key={`shared-loading-${index}`}
                className="flex items-center gap-2 rounded-md px-2 py-1.5"
              >
                <span className="size-3.5 shrink-0 animate-pulse rounded bg-muted" />
                <span
                  className="h-3.5 animate-pulse rounded bg-muted"
                  style={{ width: `${70 - index * 10}%` }}
                />
              </div>
            ))
          : items.slice(0, 6).map((item) => (
              <button
                key={item.id}
                type="button"
                onClick={() => profile.onOpenConversation?.(item.id)}
                className="flex w-full items-center gap-2 rounded-md px-2 py-1.5 text-left transition-colors hover:bg-accent/50"
              >
                <SharedConversationIcon kind={item.kind} />
                <p className="min-w-0 truncate text-[13px] font-medium text-foreground">
                  {item.title}
                </p>
              </button>
            ))}
      </div>
    </div>
  );
}

export function ProfileSidebar({
  profile,
  onClose,
  closeLabel = "Close profile",
  className,
  role,
  onClick,
  onKeyDown,
}: {
  profile: ProfileData;
  onClose: () => void;
  closeLabel?: string;
  className?: string;
  role?: "document";
  onClick?: MouseEventHandler<HTMLElement>;
  onKeyDown?: KeyboardEventHandler<HTMLElement>;
}) {
  return (
    <aside
      role={role}
      onClick={onClick}
      onKeyDown={onKeyDown}
      className={cn(
        "flex w-80 shrink-0 flex-col overflow-hidden border-l border-border bg-background",
        className,
      )}
    >
      <div className="flex items-center gap-2 border-b border-border px-4 py-3">
        <h3 className="min-w-0 flex-1 truncate text-[15px] font-bold">
          {profile.displayName}&apos;s profile
        </h3>
        <button
          type="button"
          onClick={onClose}
          className="flex size-7 items-center justify-center rounded text-muted-foreground hover:bg-accent hover:text-foreground"
          aria-label={closeLabel}
        >
          <X className="size-4" />
        </button>
      </div>

      <div className="flex-1 overflow-y-auto">
        <div className="flex flex-col items-center gap-3 px-6 pt-8 pb-6 text-center">
          <div className="shrink-0 rounded-full ring-4 ring-background">
            <ProfileAvatar
              src={profile.avatarFullSrc ?? profile.avatarThumbSrc}
              name={profile.displayName}
              size="xl"
              presence={profile.presence}
              fallbackReady={profile.avatarFallbackReady}
            />
          </div>
          <div className="min-w-0 max-w-full space-y-1 text-center">
            {profile.email ? (
              <a
                href={`mailto:${profile.email}`}
                className="block text-[13px] font-medium break-all text-primary hover:underline"
                onClick={(e) => {
                  e.preventDefault();
                  import(
                    "@better-teams/app/services/desktop/open-external"
                  ).then(({ openExternal }) =>
                    openExternal(`mailto:${profile.email}`),
                  );
                }}
              >
                {profile.email}
              </a>
            ) : null}
            {profile.jobTitle ? (
              <div className="text-[13px] leading-snug text-muted-foreground">
                {profile.jobTitle}
              </div>
            ) : null}
          </div>
        </div>

        <div className="mx-6 border-t border-border" />

        <SharedConversationList profile={profile} />
      </div>

      {profile.onMessage ? (
        <div className="flex shrink-0 gap-2 border-t border-border px-4 py-3">
          <button
            type="button"
            onClick={profile.onMessage}
            className="flex flex-1 items-center justify-center gap-2 rounded-lg border border-border px-3 py-2 text-[13px] font-medium text-foreground transition-colors hover:bg-accent"
          >
            <MessageSquare className="size-4" />
            Message
          </button>
        </div>
      ) : null}
    </aside>
  );
}

export function MembersSidebar({
  title = "Members",
  members,
  memberCount,
  onOpenProfile,
  onClose,
  closeLabel = "Close members",
  className,
}: {
  title?: string;
  members: ProfileData[];
  memberCount: number | null;
  onOpenProfile?: (profile: ProfileData) => void;
  onClose: () => void;
  closeLabel?: string;
  className?: string;
}) {
  return (
    <aside
      className={cn(
        "flex w-80 shrink-0 flex-col overflow-hidden border-l border-border bg-background",
        className,
      )}
    >
      <div className="flex items-center gap-2 border-b border-border px-4 py-3">
        <div className="min-w-0 flex-1">
          <h3 className="truncate text-[15px] font-bold">{title}</h3>
          {memberCount != null ? (
            <p className="text-[12px] tabular-nums text-muted-foreground">
              {memberCount} members
            </p>
          ) : null}
        </div>
        <button
          type="button"
          onClick={onClose}
          className="flex size-7 items-center justify-center rounded text-muted-foreground hover:bg-accent hover:text-foreground"
          aria-label={closeLabel}
        >
          <X className="size-4" />
        </button>
      </div>

      <div className="flex-1 overflow-y-auto px-3 py-3">
        {members.length > 0 ? (
          <div className="space-y-1">
            {members.map((member) => (
              <button
                key={member.mri}
                type="button"
                onClick={() => onOpenProfile?.(member)}
                className="flex w-full items-center gap-3 rounded-lg px-3 py-2 text-left transition-colors hover:bg-accent/55"
                aria-label={`View profile for ${member.displayName}`}
              >
                <ProfileAvatar
                  src={member.avatarThumbSrc}
                  name={member.displayName}
                  size="sm"
                  presence={member.presence}
                  fallbackReady={member.avatarFallbackReady}
                />
                <div className="min-w-0 flex-1">
                  <p className="truncate text-[13px] font-semibold leading-tight text-foreground">
                    {member.displayName}
                  </p>
                  {member.jobTitle || member.email ? (
                    <p className="truncate pt-0.5 text-[12px] leading-tight text-muted-foreground">
                      {member.jobTitle ?? member.email}
                    </p>
                  ) : null}
                </div>
              </button>
            ))}
          </div>
        ) : (
          <p className="px-3 py-8 text-center text-[13px] text-muted-foreground/60">
            No members loaded yet.
          </p>
        )}
      </div>
    </aside>
  );
}

export function ProfileTrigger({
  profile,
  children,
  onOpenProfile,
}: {
  profile: ProfileData | null;
  children: ReactNode;
  onOpenProfile?: (profile: ProfileData) => void;
}) {
  const [showCard, setShowCard] = useState(false);
  const [hoverTimer, setHoverTimer] = useState<ReturnType<
    typeof setTimeout
  > | null>(null);

  if (!profile) return <>{children}</>;

  const startHover = () => {
    const timer = setTimeout(() => setShowCard(true), 400);
    setHoverTimer(timer);
  };

  const endHover = () => {
    if (hoverTimer) {
      clearTimeout(hoverTimer);
      setHoverTimer(null);
    }
    setShowCard(false);
  };

  const handleClick = () => {
    setShowCard(false);
    if (hoverTimer) {
      clearTimeout(hoverTimer);
      setHoverTimer(null);
    }
    onOpenProfile?.(profile);
  };

  return (
    <button
      type="button"
      className="relative inline-flex cursor-pointer appearance-none border-none bg-transparent p-0 text-left outline-none"
      onPointerEnter={startHover}
      onPointerLeave={endHover}
      onClick={handleClick}
      aria-label={`View profile: ${profile.displayName}`}
    >
      {children}
      {showCard ? (
        <span
          className={cn(
            "absolute left-0 top-full z-40 mt-2 w-72 overflow-hidden rounded-2xl border border-border/80 bg-background/95 p-0 shadow-xl backdrop-blur",
            "animate-in fade-in-0 zoom-in-95 duration-150",
          )}
        >
          <div className="border-b border-border/70 bg-gradient-to-b from-accent/60 to-background px-3 py-3">
            <HoverProfileCard profile={profile} />
          </div>
          <div className="space-y-2 px-3 py-2.5">
            {profile.email ? (
              <div className="truncate text-[12px] text-muted-foreground">
                {profile.email}
              </div>
            ) : null}
            <div className="text-[11px] font-medium text-muted-foreground/70">
              Click to open full profile
            </div>
          </div>
        </span>
      ) : null}
    </button>
  );
}
