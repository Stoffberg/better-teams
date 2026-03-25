import { ChevronRight, Hash, Video } from "lucide-react";
import React, { useState } from "react";
import type { ConversationChatKind } from "@/lib/chat-format";
import {
  conversationKindShortLabel,
  initialsFromLabel,
} from "@/lib/chat-format";
import { cn } from "@/lib/utils";
import type { SidebarConversationItem } from "./types";

function Glyph({
  kind,
  title,
  avatarSrc,
}: {
  kind: ConversationChatKind;
  title: string;
  avatarSrc?: string;
}) {
  const [imgFailed, setImgFailed] = useState(false);

  if (kind === "dm") {
    if (avatarSrc && !imgFailed) {
      return (
        <img
          src={avatarSrc}
          alt=""
          onError={() => setImgFailed(true)}
          className="size-5 shrink-0 rounded-full object-cover"
        />
      );
    }
    return (
      <span className="flex size-5 shrink-0 items-center justify-center rounded-full bg-accent text-[7px] font-medium text-muted-foreground">
        {initialsFromLabel(title)}
      </span>
    );
  }
  if (kind === "group") {
    return (
      <span className="flex size-5 shrink-0 items-center justify-center rounded bg-sidebar-accent">
        <Hash
          className="size-2.5 stroke-[2.5] text-sidebar-foreground/40"
          aria-hidden
        />
      </span>
    );
  }
  return (
    <span className="flex size-5 shrink-0 items-center justify-center rounded bg-sidebar-accent">
      <Video
        className="size-2.5 text-sidebar-foreground/40"
        strokeWidth={2}
        aria-hidden
      />
    </span>
  );
}

export const SidebarSection = React.memo(function SidebarSection({
  title,
  count,
  items,
  open,
  onOpenChange,
  selectedId,
  onSelect,
  onConversationKeyDown,
  schedulePrefetch,
  cancelPrefetchSchedule,
  registerConversationRef,
}: {
  title: string;
  count: number;
  items: SidebarConversationItem[];
  open: boolean;
  onOpenChange: (value: boolean) => void;
  selectedId: string | null;
  onSelect: (id: string) => void;
  onConversationKeyDown: (
    event: React.KeyboardEvent<HTMLButtonElement>,
    id: string,
  ) => void;
  schedulePrefetch: (id: string, mode?: "delayed" | "immediate") => void;
  cancelPrefetchSchedule: () => void;
  registerConversationRef: (id: string, node: HTMLButtonElement | null) => void;
}) {
  if (items.length === 0) return null;

  return (
    <div>
      <button
        type="button"
        onClick={() => onOpenChange(!open)}
        className="flex h-7 w-full items-center gap-1 px-2 text-[10px] font-semibold tracking-wider text-sidebar-foreground/40 uppercase transition-colors hover:text-sidebar-foreground/60"
        aria-label={title}
        aria-expanded={open}
      >
        <ChevronRight
          className={cn(
            "size-2.5 shrink-0 transition-transform duration-150",
            open && "rotate-90",
          )}
        />
        <span>{title}</span>
        <span className="ml-auto font-normal tabular-nums">{count}</span>
      </button>
      {open ? (
        <div className="flex flex-col pb-1.5">
          {items.map((item, index) => {
            const active = item.id === selectedId;
            return (
              <button
                key={item.id}
                ref={(node) => registerConversationRef(item.id, node)}
                data-conversation-id={item.id}
                type="button"
                aria-current={active ? "true" : undefined}
                tabIndex={active || (!selectedId && index === 0) ? 0 : -1}
                aria-label={`${item.title}, ${conversationKindShortLabel(item.kind)}`}
                onClick={() => onSelect(item.id)}
                onKeyDown={(event) => onConversationKeyDown(event, item.id)}
                onPointerDown={() => schedulePrefetch(item.id, "immediate")}
                onPointerEnter={() => schedulePrefetch(item.id)}
                onPointerLeave={cancelPrefetchSchedule}
                onFocus={() => schedulePrefetch(item.id, "immediate")}
                className={cn(
                  "flex h-[30px] w-full items-center gap-2 rounded-md px-2 text-left text-[13px] outline-none transition-colors duration-75",
                  active
                    ? "bg-sidebar-selection text-sidebar-foreground"
                    : "text-sidebar-foreground/80 hover:bg-sidebar-accent/70 hover:text-sidebar-foreground",
                )}
              >
                <Glyph
                  kind={item.kind}
                  title={item.title}
                  avatarSrc={item.avatarThumbSrc}
                />
                <span className="min-w-0 flex-1 truncate">{item.title}</span>
                {item.sideTime ? (
                  <span className="shrink-0 text-[9px] tabular-nums text-sidebar-foreground/30">
                    {item.sideTime}
                  </span>
                ) : null}
              </button>
            );
          })}
        </div>
      ) : null}
    </div>
  );
});
