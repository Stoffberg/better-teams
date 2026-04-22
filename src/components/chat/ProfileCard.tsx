import {
  Briefcase,
  Building2,
  Hash,
  Mail,
  MapPin,
  MessageSquare,
  Shapes,
  Video,
  X,
} from "lucide-react";
import {
  type ComponentType,
  type KeyboardEventHandler,
  type MouseEventHandler,
  type ReactNode,
  useState,
} from "react";
import { createPortal } from "react-dom";
import { initialsFromLabel } from "@/lib/chat-format";
import { presenceDescription } from "@/lib/teams-presence";
import { cn } from "@/lib/utils";
import type { PresenceInfo } from "@/services/teams/types";
import { PresenceBadge } from "./PresenceBadge";

export type ProfileData = {
  mri: string;
  displayName: string;
  avatarThumbSrc?: string;
  avatarFullSrc?: string;
  email?: string;
  jobTitle?: string;
  department?: string;
  companyName?: string;
  tenantName?: string;
  location?: string;
  presence?: PresenceInfo | null;
  currentConversationId?: string;
  sharedConversationHeading?: string;
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
}: {
  src?: string;
  name: string;
  size: "sm" | "lg" | "xl";
  presence?: PresenceInfo | null;
}) {
  const [failed, setFailed] = useState(false);
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
  return (
    <span className="relative block">
      <span
        className={cn(
          "flex shrink-0 items-center justify-center bg-primary/10 text-primary shadow-md",
          cls,
          radius,
          textCls,
        )}
      >
        {initialsFromLabel(name)}
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

function ProfileInfoRow({
  icon: Icon,
  label,
  isLink,
  href,
}: {
  icon: ComponentType<{ className?: string }>;
  label: string;
  isLink?: boolean;
  href?: string;
}) {
  const content = (
    <div className="flex items-center gap-3 py-2">
      <Icon className="size-4 shrink-0 text-muted-foreground/60" />
      {isLink && href ? (
        <a
          href={href}
          className="min-w-0 cursor-pointer truncate text-[14px] text-primary hover:underline"
          onClick={(e) => {
            e.preventDefault();
            import("@/lib/open-external").then(({ openExternal }) =>
              openExternal(href),
            );
          }}
        >
          {label}
        </a>
      ) : (
        <span className="min-w-0 truncate text-[14px] text-foreground">
          {label}
        </span>
      )}
    </div>
  );
  return content;
}

function ProfileFact({
  icon: Icon,
  label,
  value,
}: {
  icon: ComponentType<{ className?: string }>;
  label: string;
  value: string;
}) {
  return (
    <div className="rounded-xl border border-border/80 bg-accent/35 px-3 py-3">
      <div className="flex items-start gap-3">
        <div className="mt-0.5 flex size-8 shrink-0 items-center justify-center rounded-lg bg-background">
          <Icon className="size-4 text-muted-foreground/70" />
        </div>
        <div className="min-w-0">
          <p className="text-[11px] font-medium tracking-wide text-muted-foreground uppercase">
            {label}
          </p>
          <p className="pt-0.5 text-[13px] font-medium leading-snug text-foreground">
            {value}
          </p>
        </div>
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
  if (items.length === 0) return null;

  return (
    <div className="px-6 pt-5">
      <div className="mb-1">
        <h5 className="text-[11px] font-semibold tracking-wide text-muted-foreground uppercase">
          OTHER CHATS
        </h5>
      </div>
      <div className="max-h-56 space-y-0.5 overflow-y-auto pr-1">
        {items.slice(0, 6).map((item) => (
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

function profileFacts(profile: ProfileData) {
  const orgName = profileOrgName(profile);
  return [
    orgName ? { icon: Building2, label: "Organization", value: orgName } : null,
    profile.department
      ? { icon: Shapes, label: "Department", value: profile.department }
      : null,
    profile.location
      ? { icon: MapPin, label: "Location", value: profile.location }
      : null,
  ].filter(Boolean) as Array<{
    icon: ComponentType<{ className?: string }>;
    label: string;
    value: string;
  }>;
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
  const orgName = profileOrgName(profile);
  const facts = profileFacts(profile);

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
        <div className="flex flex-col items-center gap-3 px-6 pt-8 pb-6">
          <div className="rounded-full ring-4 ring-background">
            <ProfileAvatar
              src={profile.avatarFullSrc ?? profile.avatarThumbSrc}
              name={profile.displayName}
              size="xl"
              presence={profile.presence}
            />
          </div>
          <div className="space-y-0.5 text-center">
            <h4 className="text-[20px] font-bold leading-snug tracking-[-0.01em]">
              {profile.displayName}
            </h4>
            {profile.jobTitle ? (
              <p className="text-[14px] text-muted-foreground">
                {profile.jobTitle}
              </p>
            ) : null}
            {profile.presence ? (
              <p className="text-[13px] text-muted-foreground">
                {presenceDescription(profile.presence)}
              </p>
            ) : null}
            {orgName ? (
              <p className="text-[13px] font-medium text-muted-foreground/90">
                {orgName}
              </p>
            ) : null}
          </div>
        </div>

        <div className="mx-6 border-t border-border" />

        {facts.length > 0 ? (
          <div className="grid gap-2 px-6 pt-5">
            {facts.map((fact) => (
              <ProfileFact
                key={fact.label}
                icon={fact.icon}
                label={fact.label}
                value={fact.value}
              />
            ))}
          </div>
        ) : null}

        <SharedConversationList profile={profile} />

        <div className="space-y-0 px-6 pt-5 pb-6">
          {profile.email ? (
            <ProfileInfoRow
              icon={Mail}
              label={profile.email}
              isLink
              href={`mailto:${profile.email}`}
            />
          ) : null}

          {profile.jobTitle ? (
            <ProfileInfoRow icon={Briefcase} label={profile.jobTitle} />
          ) : null}
        </div>
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

export function ProfileTrigger({
  profile,
  children,
}: {
  profile: ProfileData | null;
  children: ReactNode;
}) {
  const [showCard, setShowCard] = useState(false);
  const [showModal, setShowModal] = useState(false);
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
    setShowModal(true);
  };

  return (
    <>
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
      {showModal
        ? createPortal(
            <div
              role="dialog"
              aria-modal="true"
              aria-label={`Profile: ${profile.displayName}`}
              className="fixed inset-0 z-50 flex justify-end"
              onClick={() => setShowModal(false)}
              onKeyDown={(e) => {
                if (e.key === "Escape") setShowModal(false);
              }}
            >
              <div
                className="fixed inset-0 bg-foreground/10 backdrop-blur-[3px] animate-in fade-in-0 duration-200"
                aria-hidden="true"
              />
              <ProfileSidebar
                profile={profile}
                onClose={() => setShowModal(false)}
                closeLabel="Close profile"
                role="document"
                className="relative z-10 h-full w-full max-w-sm shadow-xl animate-in slide-in-from-right duration-200"
                onClick={(e) => e.stopPropagation()}
                onKeyDown={(e) => e.stopPropagation()}
              />
            </div>,
            document.body,
          )
        : null}
    </>
  );
}
