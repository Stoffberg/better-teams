import { ChevronDown, Hash, Search, Users, Video, X } from "lucide-react";
import {
  startTransition,
  useDeferredValue,
  useEffect,
  useRef,
  useState,
} from "react";
import type { ConversationChatKind } from "@/lib/chat-format";
import { canonAvatarMri } from "@/lib/teams-profile-avatars";
import type { PresenceInfo } from "@/services/teams/types";
import { PresenceBadge } from "./PresenceBadge";

function kindLabel(kind: ConversationChatKind): string {
  if (kind === "group") return "Group";
  if (kind === "meeting") return "Meeting";
  return "Direct";
}

function KindIcon({ kind }: { kind: ConversationChatKind }) {
  if (kind === "group") {
    return <Hash className="size-5 stroke-[2.5] text-foreground/60" />;
  }
  if (kind === "meeting") {
    return <Video className="size-5 text-foreground/60" />;
  }
  return null;
}

export function ThreadHeader({
  title,
  kind,
  memberCount,
  avatarMris,
  avatarByMri,
  presenceByMri,
  onOpenProfile,
  profileButtonLabel,
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
  presenceByMri: Record<string, PresenceInfo>;
  onOpenProfile?: () => void;
  profileButtonLabel?: string;
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
    <>
      <h2 className="min-w-0 truncate text-[16px] font-bold leading-tight tracking-[-0.01em]">
        {title}
      </h2>
      {onOpenProfile ? (
        <ChevronDown className="size-3.5 shrink-0 text-muted-foreground" />
      ) : null}
      <span className="rounded-full bg-accent px-2 py-0.5 text-[11px] font-semibold text-muted-foreground">
        {kindLabel(kind)}
      </span>
    </>
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
          <>
            <div className="flex -space-x-1.5">
              {avatarMris.slice(0, 3).map((mri) => {
                const normalizedMri = canonAvatarMri(mri);
                const avatarSrc = avatarByMri[normalizedMri];
                const presence = presenceByMri[normalizedMri];
                return avatarSrc ? (
                  <div key={mri} className="relative">
                    <img
                      src={avatarSrc}
                      alt=""
                      className="size-6 rounded-full border-2 border-background object-cover"
                    />
                    <PresenceBadge
                      presence={presence}
                      size="sm"
                      className="ring-background"
                    />
                  </div>
                ) : (
                  <div key={`placeholder-${mri}`} className="relative">
                    <div className="size-6 rounded-full border-2 border-background bg-accent" />
                    <PresenceBadge
                      presence={presence}
                      size="sm"
                      className="ring-background"
                    />
                  </div>
                );
              })}
            </div>
            {memberCount != null ? (
              <span className="inline-flex items-center gap-1 text-[13px] tabular-nums text-muted-foreground">
                <Users className="size-3.5" />
                {memberCount}
              </span>
            ) : null}
          </>
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
