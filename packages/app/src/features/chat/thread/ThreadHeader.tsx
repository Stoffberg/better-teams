import type { ConversationChatKind } from "@better-teams/core/chat";
import { canonAvatarMri } from "@better-teams/core/teams/profile/avatars";
import type { PresenceInfo } from "@better-teams/core/teams/types";
import { Hash, Search, Users, Video, X } from "lucide-react";
import {
  startTransition,
  useDeferredValue,
  useEffect,
  useRef,
  useState,
} from "react";
import { avatarFallbackPresentation } from "../avatar/avatar-fallback";
import { PresenceBadge } from "../presence/PresenceBadge";

function KindIcon({ kind }: { kind: ConversationChatKind }) {
  if (kind === "group") {
    return <Hash className="size-5 stroke-[2.5] text-foreground/60" />;
  }
  if (kind === "meeting") {
    return <Video className="size-5 text-foreground/60" />;
  }
  return null;
}

function HeaderParticipantAvatar({
  src,
  label,
  presence,
  fallbackReady,
}: {
  src?: string;
  label: string;
  presence?: PresenceInfo;
  fallbackReady: boolean;
}) {
  const [failed, setFailed] = useState(false);
  const fallback = avatarFallbackPresentation(label);

  if (src && !failed) {
    return (
      <div className="relative">
        <img
          src={src}
          alt=""
          onError={() => setFailed(true)}
          className="size-6 rounded-full border-2 border-background object-cover"
        />
        <PresenceBadge
          presence={presence}
          size="sm"
          className="size-2 ring-background"
        />
      </div>
    );
  }
  if (!fallbackReady && !failed) {
    return (
      <div
        className="size-6 rounded-full border-2 border-background"
        aria-hidden
      />
    );
  }
  return (
    <div className="relative">
      <div
        className="flex size-6 items-center justify-center rounded-full border-2 border-background text-[9px] font-semibold"
        style={fallback.style}
      >
        {fallback.initials}
      </div>
      <PresenceBadge
        presence={presence}
        size="sm"
        className="size-2 ring-background"
      />
    </div>
  );
}

export function ThreadHeader({
  title,
  kind,
  memberCount,
  avatarMris,
  avatarByMri,
  avatarLabelByMri,
  avatarFallbackReady = true,
  presenceByMri,
  onOpenProfile,
  profileButtonLabel,
  onOpenMembers,
  searchQuery,
  searchResultCount,
  onSearchQueryChange,
  onSubmitSearch,
  onCloseSearch,
}: {
  title: string;
  kind: ConversationChatKind;
  memberCount: number | null;
  avatarMris: string[];
  avatarByMri: Record<string, string>;
  avatarLabelByMri: Record<string, string>;
  avatarFallbackReady?: boolean;
  presenceByMri: Record<string, PresenceInfo>;
  onOpenProfile?: () => void;
  profileButtonLabel?: string;
  onOpenMembers?: () => void;
  searchQuery: string;
  searchResultCount: number;
  onSearchQueryChange: (value: string) => void;
  onSubmitSearch: (query: string) => void;
  onCloseSearch: () => void;
}) {
  const [showSearch, setShowSearch] = useState(false);
  const [searchDraft, setSearchDraft] = useState(searchQuery);
  const deferredSearchDraft = useDeferredValue(searchDraft);
  const searchInputRef = useRef<HTMLInputElement>(null);
  const showParticipantSummary = kind !== "dm";

  useEffect(() => {
    if (!showSearch) return;
    searchInputRef.current?.focus();
    searchInputRef.current?.select();
  }, [showSearch]);

  useEffect(() => {
    setSearchDraft(searchQuery);
  }, [searchQuery]);

  useEffect(() => {
    const trimmedDraft = deferredSearchDraft.trim();
    if (trimmedDraft === searchQuery.trim()) return;
    const timer = window.setTimeout(() => {
      startTransition(() => {
        onSearchQueryChange(deferredSearchDraft);
      });
    }, 140);
    return () => window.clearTimeout(timer);
  }, [deferredSearchDraft, onSearchQueryChange, searchQuery]);

  const trimmedSearchDraft = deferredSearchDraft.trim();
  const searchSummary = !trimmedSearchDraft
    ? null
    : trimmedSearchDraft !== searchQuery.trim()
      ? "Searching"
      : searchResultCount === 1
        ? "1 result"
        : `${searchResultCount} results`;

  const titleBlock = (
    <h2 className="min-w-0 truncate text-[16px] font-bold leading-tight tracking-[-0.01em]">
      {title}
    </h2>
  );

  return (
    <header className="flex shrink-0 items-center gap-3 border-b border-border bg-background px-5 py-2.5">
      <div className="flex min-w-0 flex-1 items-center gap-1.5">
        <KindIcon kind={kind} />
        {onOpenProfile ? (
          <button
            type="button"
            onClick={onOpenProfile}
            className="flex min-w-0 items-center gap-2 rounded-md px-1 py-1 text-left transition-colors hover:bg-accent"
            aria-label={profileButtonLabel ?? `View profile for ${title}`}
          >
            {titleBlock}
          </button>
        ) : (
          <div className="flex min-w-0 items-center gap-2">{titleBlock}</div>
        )}
      </div>

      <div className="flex items-center gap-3">
        {showSearch ? (
          <div className="flex items-center gap-1 rounded-lg border border-border bg-accent/40 px-2 py-1">
            <Search className="size-3.5 text-muted-foreground" />
            <input
              ref={searchInputRef}
              type="search"
              value={searchDraft}
              onChange={(event) => {
                const nextValue = event.target.value;
                startTransition(() => {
                  setSearchDraft(nextValue);
                });
              }}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  onSubmitSearch(searchDraft);
                }
                if (event.key === "Escape") {
                  setShowSearch(false);
                  setSearchDraft("");
                  onCloseSearch();
                }
              }}
              placeholder="Find in chat"
              aria-label="Find in conversation"
              className="w-32 bg-transparent text-[12px] text-foreground outline-none placeholder:text-muted-foreground"
            />
            {searchSummary ? (
              <span
                className="shrink-0 text-[11px] tabular-nums text-muted-foreground"
                aria-live="polite"
              >
                {searchSummary}
              </span>
            ) : null}
            <button
              type="button"
              onClick={() => {
                setShowSearch(false);
                setSearchDraft("");
                onCloseSearch();
              }}
              className="text-muted-foreground transition-colors hover:text-foreground"
              aria-label="Close search"
            >
              <X className="size-3.5" />
            </button>
          </div>
        ) : null}
        {showParticipantSummary ? (
          <button
            type="button"
            onClick={onOpenMembers}
            className="flex items-center gap-2 rounded-lg px-1.5 py-1 transition-colors hover:bg-accent"
            aria-label={
              memberCount != null
                ? `Open members (${memberCount})`
                : "Open members"
            }
          >
            {avatarMris.length > 0 ? (
              <div className="flex -space-x-1.5">
                {avatarMris.slice(0, 3).map((mri) => {
                  const normalizedMri = canonAvatarMri(mri);
                  const avatarSrc = avatarByMri[normalizedMri];
                  const label = avatarLabelByMri[normalizedMri] ?? "";
                  const presence = presenceByMri[normalizedMri];
                  return (
                    <HeaderParticipantAvatar
                      key={mri}
                      src={avatarSrc}
                      label={label}
                      presence={presence}
                      fallbackReady={avatarFallbackReady}
                    />
                  );
                })}
              </div>
            ) : null}
            {memberCount != null ? (
              <span className="inline-flex items-center gap-1 text-[13px] tabular-nums text-muted-foreground">
                <Users className="size-3.5" />
                {memberCount}
              </span>
            ) : null}
          </button>
        ) : null}
        <button
          type="button"
          onClick={() => setShowSearch((current) => !current)}
          className="flex size-8 items-center justify-center rounded-lg text-muted-foreground/60 transition-colors hover:bg-accent hover:text-foreground"
          aria-label="Search in conversation"
        >
          <Search className="size-4" />
        </button>
      </div>
    </header>
  );
}
