import { initialsFromLabel } from "@better-teams/core/chat";
import type {
  PresenceInfo,
  TeamsAccountOption,
} from "@better-teams/core/teams/types";
import {
  Avatar,
  AvatarFallback,
  AvatarImage,
} from "@better-teams/ui/components/avatar";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@better-teams/ui/components/dropdown-menu";
import { Skeleton } from "@better-teams/ui/components/skeleton";
import { cn } from "@better-teams/ui/utils";
import {
  Check,
  ChevronDown,
  Search,
  Star,
  Users,
  Video,
  X,
} from "lucide-react";
import React, {
  useCallback,
  useDeferredValue,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { PresenceBadge } from "./PresenceBadge";
import type { SidebarConversationItem } from "./types";

function workspaceLabelFromUpn(upn?: string): string {
  const email = upn?.trim().toLowerCase();
  if (!email || !email.includes("@")) return "Workspace";
  const domain = email.split("@")[1] ?? "";
  const root = domain.split(".")[0] ?? "";
  if (!root) return "Workspace";
  return root
    .split(/[-_]+/)
    .filter(Boolean)
    .map((part) => part[0]?.toUpperCase() + part.slice(1))
    .join(" ");
}

function conversationAriaLabel(item: SidebarConversationItem): string {
  const kindLabel =
    item.kind === "dm"
      ? "direct message"
      : item.kind === "group"
        ? "group chat"
        : "meeting";
  return `${item.title}, ${kindLabel}`;
}

function favoriteButtonLabel(item: SidebarConversationItem): string {
  return item.isFavorite
    ? `Remove ${item.title} from favorites`
    : `Add ${item.title} to favorites`;
}

function SidebarHeaderSkeleton() {
  return (
    <div className="flex w-full items-center gap-3 rounded-xl px-2 py-2">
      <Skeleton className="size-10 shrink-0 rounded-xl bg-sidebar-accent" />
      <span className="min-w-0 flex-1 space-y-2">
        <Skeleton className="h-3.5 w-24 bg-sidebar-accent" />
        <Skeleton className="h-3 w-36 bg-sidebar-accent" />
      </span>
    </div>
  );
}

const SIDEBAR_SKELETON_ROWS = [
  { key: "a", width: "72%" },
  { key: "b", width: "92%" },
  { key: "c", width: "64%" },
  { key: "d", width: "84%" },
  { key: "e", width: "76%" },
  { key: "f", width: "58%" },
  { key: "g", width: "88%" },
  { key: "h", width: "68%" },
];

function SidebarListSkeleton() {
  return (
    <div className="space-y-1.5 px-1">
      {SIDEBAR_SKELETON_ROWS.map((row) => (
        <div
          key={row.key}
          className="flex items-center gap-2.5 rounded-md px-2 py-1.5"
        >
          <Skeleton className="size-5 shrink-0 rounded-md bg-sidebar-accent" />
          <Skeleton
            className="h-3.5 bg-sidebar-accent"
            style={{ width: row.width }}
          />
        </div>
      ))}
    </div>
  );
}

function ConversationRow({
  item,
  active,
  onClick,
  onHoverStart,
  onHoverEnd,
  onKeyDown,
  registerRef,
  tabbable,
  children,
  action,
}: {
  item: SidebarConversationItem;
  active: boolean;
  onClick: () => void;
  onHoverStart: () => void;
  onHoverEnd: () => void;
  onKeyDown: (e: React.KeyboardEvent<HTMLButtonElement>) => void;
  registerRef: (node: HTMLButtonElement | null) => void;
  tabbable: boolean;
  children: React.ReactNode;
  action?: React.ReactNode;
}) {
  return (
    <div className="group relative" data-sidebar-row>
      <button
        ref={registerRef}
        data-conversation-id={item.id}
        type="button"
        aria-current={active ? "true" : undefined}
        tabIndex={tabbable ? 0 : -1}
        aria-label={conversationAriaLabel(item)}
        onClick={onClick}
        onPointerEnter={onHoverStart}
        onPointerLeave={onHoverEnd}
        onFocus={onHoverStart}
        onBlur={onHoverEnd}
        onKeyDown={onKeyDown}
        className={cn(
          "flex min-w-0 w-full items-center gap-2.5 rounded-md px-3 py-1.5 pr-10 text-left text-[14px] outline-none transition-colors duration-75",
          active
            ? "bg-primary/10 font-bold text-sidebar-foreground"
            : "text-sidebar-foreground/70 hover:bg-sidebar-accent",
        )}
      >
        {children}
      </button>
      {action ? (
        <div className="absolute top-1/2 right-3 -translate-y-1/2">
          {action}
        </div>
      ) : null}
    </div>
  );
}

const ChannelItem = React.memo(function ChannelItem({
  item,
  active,
  onSelect,
  onHoverStart,
  onHoverEnd,
  onKeyDown,
  registerRef,
  tabbable,
  onToggleFavorite,
}: {
  item: SidebarConversationItem;
  active: boolean;
  onSelect: (id: string, focus: "sidebar" | "thread" | "composer") => void;
  onHoverStart: (conversationId: string) => void;
  onHoverEnd: (conversationId: string) => void;
  onKeyDown: (e: React.KeyboardEvent<HTMLButtonElement>, id: string) => void;
  registerRef: React.RefObject<Record<string, HTMLButtonElement | null>>;
  tabbable: boolean;
  onToggleFavorite: (conversationId: string, favorite: boolean) => void;
}) {
  const isMeeting = item.kind === "meeting";
  const handleClick = useCallback(
    () => onSelect(item.id, "thread"),
    [item.id, onSelect],
  );
  const handleHoverStart = useCallback(
    () => onHoverStart(item.id),
    [item.id, onHoverStart],
  );
  const handleHoverEnd = useCallback(
    () => onHoverEnd(item.id),
    [item.id, onHoverEnd],
  );
  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLButtonElement>) => onKeyDown(e, item.id),
    [item.id, onKeyDown],
  );
  const handleRegisterRef = useCallback(
    (node: HTMLButtonElement | null) => {
      registerRef.current[item.id] = node;
    },
    [item.id, registerRef],
  );

  return (
    <ConversationRow
      item={item}
      active={active}
      onClick={handleClick}
      onHoverStart={handleHoverStart}
      onHoverEnd={handleHoverEnd}
      onKeyDown={handleKeyDown}
      registerRef={handleRegisterRef}
      tabbable={tabbable}
      action={
        <button
          type="button"
          aria-label={favoriteButtonLabel(item)}
          aria-pressed={item.isFavorite}
          onClick={(event) => {
            event.stopPropagation();
            onToggleFavorite(item.id, !item.isFavorite);
          }}
          className="flex size-5 items-center justify-center rounded-md text-sidebar-muted opacity-0 transition-opacity hover:text-sidebar-foreground focus-visible:opacity-100 group-hover:opacity-100 group-focus-within:opacity-100"
        >
          <Star
            className={cn("size-3.5", item.isFavorite && "fill-current")}
            strokeWidth={1.9}
          />
        </button>
      }
    >
      <span
        className="flex size-5 shrink-0 items-center justify-center rounded-md text-sidebar-foreground/35"
        aria-hidden
      >
        {isMeeting ? (
          <Video className="size-4" strokeWidth={1.75} />
        ) : (
          <Users className="size-4" strokeWidth={1.75} />
        )}
      </span>
      <span className="min-w-0 flex-1 truncate leading-5">{item.title}</span>
    </ConversationRow>
  );
});

const DMItem = React.memo(function DMItem({
  item,
  active,
  presence,
  onSelect,
  onHoverStart,
  onHoverEnd,
  onKeyDown,
  registerRef,
  tabbable,
  onToggleFavorite,
}: {
  item: SidebarConversationItem;
  active: boolean;
  presence?: PresenceInfo | null;
  onSelect: (id: string, focus: "sidebar" | "thread" | "composer") => void;
  onHoverStart: (conversationId: string) => void;
  onHoverEnd: (conversationId: string) => void;
  onKeyDown: (e: React.KeyboardEvent<HTMLButtonElement>, id: string) => void;
  registerRef: React.RefObject<Record<string, HTMLButtonElement | null>>;
  tabbable: boolean;
  onToggleFavorite: (conversationId: string, favorite: boolean) => void;
}) {
  const [imgFailed, setImgFailed] = useState(false);
  const handleClick = useCallback(
    () => onSelect(item.id, "thread"),
    [item.id, onSelect],
  );
  const handleHoverStart = useCallback(
    () => onHoverStart(item.id),
    [item.id, onHoverStart],
  );
  const handleHoverEnd = useCallback(
    () => onHoverEnd(item.id),
    [item.id, onHoverEnd],
  );
  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLButtonElement>) => onKeyDown(e, item.id),
    [item.id, onKeyDown],
  );
  const handleRegisterRef = useCallback(
    (node: HTMLButtonElement | null) => {
      registerRef.current[item.id] = node;
    },
    [item.id, registerRef],
  );

  return (
    <ConversationRow
      item={item}
      active={active}
      onClick={handleClick}
      onHoverStart={handleHoverStart}
      onHoverEnd={handleHoverEnd}
      onKeyDown={handleKeyDown}
      registerRef={handleRegisterRef}
      tabbable={tabbable}
      action={
        <button
          type="button"
          aria-label={favoriteButtonLabel(item)}
          aria-pressed={item.isFavorite}
          onClick={(event) => {
            event.stopPropagation();
            onToggleFavorite(item.id, !item.isFavorite);
          }}
          className="flex size-5 items-center justify-center rounded-md text-sidebar-muted opacity-0 transition-opacity hover:text-sidebar-foreground focus-visible:opacity-100 group-hover:opacity-100 group-focus-within:opacity-100"
        >
          <Star
            className={cn("size-3.5", item.isFavorite && "fill-current")}
            strokeWidth={1.9}
          />
        </button>
      }
    >
      <span className="relative flex size-5 shrink-0 items-center justify-center">
        {item.avatarThumbSrc && !imgFailed ? (
          <img
            src={item.avatarThumbSrc}
            alt=""
            onError={() => setImgFailed(true)}
            className="size-5 rounded-md object-cover"
          />
        ) : (
          <span className="flex size-5 items-center justify-center rounded-md bg-sidebar-accent text-[9px] font-semibold text-sidebar-muted">
            {initialsFromLabel(item.title)}
          </span>
        )}
        <PresenceBadge
          presence={presence}
          size="sm"
          className="rounded-sm ring-sidebar"
        />
      </span>
      <span className="min-w-0 flex-1 truncate">{item.title}</span>
    </ConversationRow>
  );
});

export function Sidebar({
  upn,
  selfAvatarSrc: _selfAvatarSrc,
  accountAvatarByTenant,
  presenceByMri,
  accounts,
  activeTenantId,
  onSwitchAccount,
  switchPending,
  allSidebarItems,
  activeConversationId,
  onSelectConversation,
  onHoverConversationStart,
  onHoverConversationEnd,
  onToggleFavorite,
  searchInputRef,
  accountLoading = false,
  conversationsLoading = false,
}: {
  upn?: string;
  selfAvatarSrc?: string;
  accountAvatarByTenant: Record<string, string>;
  presenceByMri: Record<string, PresenceInfo>;
  accounts: TeamsAccountOption[];
  activeTenantId?: string;
  onSwitchAccount: (tenantId: string | null) => void;
  switchPending: boolean;
  allSidebarItems: SidebarConversationItem[];
  activeConversationId: string | null;
  onSelectConversation: (
    id: string,
    focus: "sidebar" | "thread" | "composer",
  ) => void;
  onHoverConversationStart: (conversationId: string) => void;
  onHoverConversationEnd: (conversationId: string) => void;
  onToggleFavorite: (conversationId: string, favorite: boolean) => void;
  searchInputRef: React.RefObject<HTMLInputElement | null>;
  accountLoading?: boolean;
  conversationsLoading?: boolean;
}) {
  const [query, setQuery] = useState("");
  const deferredQuery = useDeferredValue(query);
  const activeAccount = useMemo(
    () =>
      accounts.find((account) => account.tenantId === activeTenantId) ??
      accounts.find((account) => account.upn === upn) ??
      accounts[0],
    [accounts, activeTenantId, upn],
  );
  const activeWorkspaceLabel = workspaceLabelFromUpn(activeAccount?.upn ?? upn);
  const activeEmail = activeAccount?.upn ?? upn ?? "Unknown account";

  const conversationRowRefs = useRef<Record<string, HTMLButtonElement | null>>(
    {},
  );
  const activeConversationIdRef = useRef<string | null>(activeConversationId);
  const filteredItemsRef = useRef<SidebarConversationItem[]>([]);

  const filteredSidebarItems = useMemo(() => {
    const q = deferredQuery.trim().toLowerCase();
    let favoriteCount = 0;
    const items = q
      ? allSidebarItems.filter((item) => item.searchText.includes(q))
      : allSidebarItems;
    for (const item of items) {
      if (item.isFavorite) favoriteCount += 1;
    }
    return { items, favoriteCount };
  }, [allSidebarItems, deferredQuery]);
  activeConversationIdRef.current = activeConversationId;
  filteredItemsRef.current = filteredSidebarItems.items;

  const openConversation = useCallback(
    (
      id: string,
      focusTarget: "sidebar" | "thread" | "composer" = "sidebar",
    ) => {
      onSelectConversation(id, focusTarget);
    },
    [onSelectConversation],
  );

  const moveConversationSelection = useCallback(
    (direction: "next" | "prev") => {
      const ids = filteredItemsRef.current.map((i) => i.id);
      if (ids.length === 0) return;
      const currentId = activeConversationIdRef.current;
      const cur = currentId ? ids.indexOf(currentId) : -1;
      let next = 0;
      if (cur < 0) next = direction === "prev" ? ids.length - 1 : 0;
      else if (direction === "next") next = Math.min(cur + 1, ids.length - 1);
      else next = Math.max(cur - 1, 0);
      openConversation(ids[next], "sidebar");
    },
    [openConversation],
  );

  const onConversationKeyDown = useCallback(
    (event: React.KeyboardEvent<HTMLButtonElement>, id: string) => {
      if (event.key === "ArrowDown" || event.key.toLowerCase() === "j") {
        event.preventDefault();
        moveConversationSelection("next");
      }
      if (event.key === "ArrowUp" || event.key.toLowerCase() === "k") {
        event.preventDefault();
        moveConversationSelection("prev");
      }
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        openConversation(id, "thread");
      }
    },
    [moveConversationSelection, openConversation],
  );

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "k") {
        event.preventDefault();
        searchInputRef.current?.focus();
        searchInputRef.current?.select();
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [searchInputRef]);

  return (
    <aside
      className="flex w-64 shrink-0 flex-col overflow-hidden border-r border-sidebar-border bg-sidebar"
      aria-label="Conversations"
      onKeyDownCapture={(event) => {
        const target = event.target;
        const editable =
          target instanceof HTMLElement &&
          (target.isContentEditable ||
            target.tagName === "INPUT" ||
            target.tagName === "TEXTAREA" ||
            target.tagName === "SELECT");

        if (event.key === "Escape" && query) {
          event.preventDefault();
          setQuery("");
          return;
        }
        if (
          (event.metaKey || event.ctrlKey) &&
          event.key.toLowerCase() === "k"
        ) {
          event.preventDefault();
          searchInputRef.current?.focus();
          searchInputRef.current?.select();
          return;
        }
        if (editable) return;
        if (event.key.toLowerCase() === "j") {
          event.preventDefault();
          moveConversationSelection("next");
          return;
        }
        if (event.key.toLowerCase() === "k") {
          event.preventDefault();
          moveConversationSelection("prev");
        }
      }}
    >
      {/* Workspace header */}
      <div className="border-b border-sidebar-border px-3 py-3">
        {accountLoading ? (
          <SidebarHeaderSkeleton />
        ) : (
          <DropdownMenu>
            <DropdownMenuTrigger asChild>
              <button
                type="button"
                aria-label="Switch account"
                disabled={switchPending}
                className="flex w-full items-center gap-3 rounded-xl px-2 py-2 text-left transition-colors hover:bg-sidebar-accent disabled:opacity-60"
              >
                <Avatar className="size-10 border border-sidebar-border bg-sidebar-accent">
                  <AvatarImage src={_selfAvatarSrc} alt={activeEmail} />
                  <AvatarFallback className="bg-sidebar-accent text-[12px] font-semibold text-sidebar-foreground">
                    {initialsFromLabel(activeWorkspaceLabel)}
                  </AvatarFallback>
                </Avatar>
                <span className="min-w-0 flex-1">
                  <span className="flex items-center gap-1.5">
                    <span className="truncate text-[14px] font-semibold text-sidebar-foreground">
                      {activeWorkspaceLabel}
                    </span>
                    <ChevronDown className="size-3.5 shrink-0 text-sidebar-muted" />
                  </span>
                  <span className="block truncate pt-0.5 text-[12px] text-sidebar-muted">
                    {activeEmail}
                  </span>
                </span>
              </button>
            </DropdownMenuTrigger>
            <DropdownMenuContent
              align="start"
              className="w-72 border-border bg-background p-1.5 text-foreground"
            >
              <DropdownMenuLabel className="px-2 text-[11px] font-semibold tracking-wider text-muted-foreground uppercase">
                Switch account
              </DropdownMenuLabel>
              <DropdownMenuSeparator className="bg-border" />
              <DropdownMenuRadioGroup
                value={activeTenantId ?? "__default__"}
                onValueChange={(value) =>
                  onSwitchAccount(value === "__default__" ? null : value)
                }
              >
                {accounts.map((account) => (
                  <DropdownMenuRadioItem
                    key={`${account.tenantId ?? "default"}:${account.upn ?? ""}`}
                    value={account.tenantId ?? "__default__"}
                    className="gap-3 rounded-lg py-2.5 pr-2.5 pl-8 focus:bg-accent focus:text-foreground"
                  >
                    <Avatar size="sm">
                      <AvatarImage
                        src={
                          account.tenantId
                            ? accountAvatarByTenant[account.tenantId]
                            : _selfAvatarSrc
                        }
                        alt={account.upn}
                      />
                      <AvatarFallback className="bg-accent text-[10px] text-muted-foreground">
                        {(account.upn?.[0] ?? "?").toUpperCase()}
                      </AvatarFallback>
                    </Avatar>
                    <span className="flex min-w-0 flex-1 flex-col">
                      <span className="truncate text-[13px] font-semibold">
                        {workspaceLabelFromUpn(account.upn)}
                      </span>
                      <span className="truncate text-[11px] text-muted-foreground">
                        {account.upn ?? "Unknown account"}
                      </span>
                    </span>
                    {account.tenantId === activeTenantId ? (
                      <Check className="size-4 text-primary" />
                    ) : null}
                  </DropdownMenuRadioItem>
                ))}
              </DropdownMenuRadioGroup>
            </DropdownMenuContent>
          </DropdownMenu>
        )}
      </div>

      {/* Global search */}
      <div className="px-3 pt-3 pb-1">
        <div className="relative">
          <Search className="pointer-events-none absolute left-2.5 top-1/2 size-3.5 -translate-y-1/2 text-sidebar-muted" />
          <input
            ref={searchInputRef}
            type="search"
            placeholder="Search"
            aria-label="Search chats"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            className="w-full rounded-md border border-sidebar-border bg-sidebar-accent/50 py-1.5 pr-3 pl-8 text-[13px] text-sidebar-foreground placeholder:text-sidebar-muted outline-none transition-colors focus:border-primary/40 focus:bg-sidebar-accent"
          />
          {query ? (
            <button
              type="button"
              onClick={() => setQuery("")}
              className="absolute right-2 top-1/2 -translate-y-1/2 text-sidebar-muted hover:text-sidebar-foreground"
              aria-label="Clear search"
            >
              <X className="size-3.5" />
            </button>
          ) : null}
        </div>
      </div>

      {/* Scrollable channel/DM list */}
      <div className="flex-1 overflow-y-auto overflow-x-hidden px-2 pt-3 pb-2">
        {conversationsLoading && allSidebarItems.length === 0 && !query ? (
          <SidebarListSkeleton />
        ) : filteredSidebarItems.items.length === 0 &&
          allSidebarItems.length > 0 ? (
          <div className="flex flex-col items-center gap-2 px-4 py-16">
            <Search className="size-6 text-sidebar-muted/40" />
            <p className="text-center text-[13px] text-sidebar-muted/60">
              No matches found
            </p>
          </div>
        ) : (
          <div className="flex flex-col">
            {filteredSidebarItems.items.map((item, index) =>
              item.kind === "dm" ? (
                <div
                  key={item.id}
                  className={cn(
                    index === filteredSidebarItems.favoriteCount &&
                      "mt-2 border-t border-sidebar-border pt-2",
                  )}
                >
                  <DMItem
                    item={item}
                    active={item.id === activeConversationId}
                    tabbable={
                      item.id === activeConversationId ||
                      (!activeConversationId && index === 0)
                    }
                    presence={
                      item.avatarMri ? presenceByMri[item.avatarMri] : undefined
                    }
                    onSelect={openConversation}
                    onHoverStart={onHoverConversationStart}
                    onHoverEnd={onHoverConversationEnd}
                    onKeyDown={onConversationKeyDown}
                    registerRef={conversationRowRefs}
                    onToggleFavorite={onToggleFavorite}
                  />
                </div>
              ) : (
                <div
                  key={item.id}
                  className={cn(
                    index === filteredSidebarItems.favoriteCount &&
                      "mt-2 border-t border-sidebar-border pt-2",
                  )}
                >
                  <ChannelItem
                    item={item}
                    active={item.id === activeConversationId}
                    tabbable={
                      item.id === activeConversationId ||
                      (!activeConversationId && index === 0)
                    }
                    onSelect={openConversation}
                    onHoverStart={onHoverConversationStart}
                    onHoverEnd={onHoverConversationEnd}
                    onKeyDown={onConversationKeyDown}
                    registerRef={conversationRowRefs}
                    onToggleFavorite={onToggleFavorite}
                  />
                </div>
              ),
            )}
          </div>
        )}
      </div>
    </aside>
  );
}
