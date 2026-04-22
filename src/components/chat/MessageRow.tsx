import React, {
  type MouseEvent,
  useCallback,
  useEffect,
  useRef,
  useState,
} from "react";
import {
  initialsFromLabel,
  type MessageAttachment,
  type MessageInlinePart,
  type MessageReference,
} from "@/lib/chat-format";
import { openExternal } from "@/lib/open-external";
import { useAuthenticatedImage } from "@/lib/use-authenticated-image";
import { cn } from "@/lib/utils";
import type { PresenceInfo } from "@/services/teams/types";
import { PresenceBadge } from "./PresenceBadge";
import type { ProfileData } from "./ProfileCard";
import { ProfileTrigger } from "./ProfileCard";
import type { DisplayMessage, ReadStatus } from "./types";

// ── External Link Handler ──

function useOpenExternal() {
  return useCallback((e: MouseEvent<HTMLAnchorElement>) => {
    e.preventDefault();
    const href = e.currentTarget.href;
    if (href) openExternal(href);
  }, []);
}

// ── File Extension Icons ──

const FILE_EXT_COLORS: Record<string, string> = {
  pdf: "bg-red-500/15 text-red-600",
  doc: "bg-blue-500/15 text-blue-600",
  docx: "bg-blue-500/15 text-blue-600",
  xls: "bg-green-500/15 text-green-600",
  xlsx: "bg-green-500/15 text-green-600",
  ppt: "bg-orange-500/15 text-orange-600",
  pptx: "bg-orange-500/15 text-orange-600",
  zip: "bg-yellow-500/15 text-yellow-700",
  rar: "bg-yellow-500/15 text-yellow-700",
  "7z": "bg-yellow-500/15 text-yellow-700",
  txt: "bg-gray-500/15 text-gray-600",
  csv: "bg-green-500/15 text-green-600",
  json: "bg-purple-500/15 text-purple-600",
  png: "bg-pink-500/15 text-pink-600",
  jpg: "bg-pink-500/15 text-pink-600",
  jpeg: "bg-pink-500/15 text-pink-600",
  gif: "bg-pink-500/15 text-pink-600",
  svg: "bg-pink-500/15 text-pink-600",
  mp4: "bg-violet-500/15 text-violet-600",
  mp3: "bg-violet-500/15 text-violet-600",
};

function fileExtLabel(attachment: MessageAttachment): string {
  const ext = attachment.fileExtension;
  if (ext && ext.length <= 5) return ext.toUpperCase();
  return "FILE";
}

function fileExtColor(attachment: MessageAttachment): string {
  const ext = attachment.fileExtension;
  return (ext && FILE_EXT_COLORS[ext]) || "bg-muted text-muted-foreground";
}

// ── Avatar ──

function MsgAvatar({
  src,
  name,
  presence,
}: {
  src?: string;
  name: string;
  presence?: PresenceInfo | null;
}) {
  const [failed, setFailed] = useState(false);
  const initials = initialsFromLabel(name);

  if (src && !failed) {
    return (
      <span className="relative block">
        <img
          src={src}
          alt=""
          onError={() => setFailed(true)}
          className="size-9 rounded-lg object-cover"
        />
        <PresenceBadge
          presence={presence}
          size="lg"
          className="rounded-md ring-background"
        />
      </span>
    );
  }
  return (
    <span className="relative flex size-9 items-center justify-center rounded-lg bg-primary/10 text-[11px] font-bold text-primary">
      {initials}
      <PresenceBadge
        presence={presence}
        size="lg"
        className="rounded-md ring-background"
      />
    </span>
  );
}

// ── Read Receipt Indicator ──

function ReadReceipt({ status }: { status?: ReadStatus }) {
  if (!status || status === "sending") return null;

  if (status === "read") {
    return (
      <span className="ml-1 inline-flex items-center text-primary">
        <svg
          width="14"
          height="14"
          viewBox="0 0 16 16"
          fill="none"
          className="text-primary"
          aria-hidden="true"
        >
          <path
            d="M1.5 8.5l3 3 7-7"
            stroke="currentColor"
            strokeWidth="1.5"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
          <path
            d="M5.5 8.5l3 3 7-7"
            stroke="currentColor"
            strokeWidth="1.5"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
      </span>
    );
  }

  if (status === "delivered") {
    return (
      <span className="ml-1 inline-flex items-center text-muted-foreground/50">
        <svg
          width="14"
          height="14"
          viewBox="0 0 16 16"
          fill="none"
          className="text-muted-foreground/50"
          aria-hidden="true"
        >
          <path
            d="M1.5 8.5l3 3 7-7"
            stroke="currentColor"
            strokeWidth="1.5"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
          <path
            d="M5.5 8.5l3 3 7-7"
            stroke="currentColor"
            strokeWidth="1.5"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
      </span>
    );
  }

  // "sent" - single check
  return (
    <span className="ml-1 inline-flex items-center text-muted-foreground/40">
      <svg
        width="14"
        height="14"
        viewBox="0 0 16 16"
        fill="none"
        className="text-muted-foreground/40"
        aria-hidden="true"
      >
        <path
          d="M3 8.5l3.5 3.5 7-7"
          stroke="currentColor"
          strokeWidth="1.5"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
      </svg>
    </span>
  );
}

function ReadReceiptDetails({
  status,
  sentAt,
  readAt,
  receiptScope,
  receiptSeenBy,
  receiptUnseenBy,
}: {
  status?: ReadStatus;
  sentAt?: string;
  readAt?: string;
  receiptScope?: "dm" | "group";
  receiptSeenBy?: DisplayMessage["receiptSeenBy"];
  receiptUnseenBy?: DisplayMessage["receiptUnseenBy"];
}) {
  if (!status || status === "sending") return null;
  const isGroupReceipt = receiptScope === "group";
  const readLabel =
    status === "read" && readAt
      ? readAt
      : status === "delivered"
        ? "Delivered, not read yet"
        : "Not read yet";
  const seenBy = receiptSeenBy ?? [];
  const unseenBy = receiptUnseenBy ?? [];

  return (
    <span className="group/receipt relative ml-1 inline-flex items-center">
      <ReadReceipt status={status} />
      <span className="pointer-events-none absolute bottom-full left-0 z-20 mb-2 min-w-52 max-w-80 translate-y-1 rounded-lg border border-border bg-background/95 px-3 py-2 text-[11px] text-foreground opacity-0 shadow-sm transition-all duration-150 group-hover/receipt:pointer-events-auto group-hover/receipt:translate-y-0 group-hover/receipt:opacity-100 group-hover/receipt:shadow-md group-hover/receipt:backdrop-blur-sm">
        <span className="flex items-center justify-between gap-3 whitespace-nowrap">
          <span className="text-muted-foreground">Sent</span>
          <span className="text-right tabular-nums">{sentAt || "Unknown"}</span>
        </span>
        {isGroupReceipt ? (
          <span className="mt-2 block max-h-64 space-y-2 overflow-y-auto pr-1">
            {seenBy.length > 0 ? (
              <span className="block">
                <span className="mb-1 block text-muted-foreground">
                  Seen by
                </span>
                <span className="block space-y-1">
                  {seenBy.map((person) => (
                    <span
                      key={person.mri}
                      className="flex items-center justify-between gap-3"
                    >
                      <span className="min-w-0 truncate">{person.name}</span>
                      <span className="shrink-0 text-right tabular-nums text-muted-foreground">
                        {person.readAt}
                      </span>
                    </span>
                  ))}
                </span>
              </span>
            ) : null}
            {unseenBy.length > 0 ? (
              <span className="block">
                <span className="mb-1 block text-muted-foreground">
                  Not seen by
                </span>
                <span className="block text-foreground/90">
                  {unseenBy.map((person) => person.name).join(", ")}
                </span>
              </span>
            ) : null}
          </span>
        ) : (
          <span className="mt-1 flex items-center justify-between gap-3 whitespace-nowrap">
            <span className="text-muted-foreground">Read</span>
            <span className="text-right tabular-nums">{readLabel}</span>
          </span>
        )}
      </span>
    </span>
  );
}

// ── Size Formatter ──

function formatAttachmentSize(size?: number): string | null {
  if (typeof size !== "number" || !Number.isFinite(size) || size <= 0) {
    return null;
  }
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  if (size < 1024 * 1024 * 1024) {
    return `${(size / (1024 * 1024)).toFixed(1)} MB`;
  }
  return `${(size / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

// ── Download Icon ──

function DownloadIcon({ className }: { className?: string }) {
  return (
    <svg
      width="16"
      height="16"
      viewBox="0 0 16 16"
      fill="none"
      className={className}
      aria-hidden="true"
    >
      <path
        d="M8 2v8m0 0L5 7m3 3l3-3M3 12h10"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

// ── Attachment Cards ──

/**
 * Build candidate image URLs to try, in priority order.
 * AMS objectUrls need a view suffix; SharePoint files use filePreview.previewUrl.
 */
function resolveImageFetchUrls(attachment: MessageAttachment): string[] {
  const urls: string[] = [];
  const base = attachment.objectUrl;
  const isAms = base && /\/v1\/objects\/[^/]+$/.test(base);
  // 1. AMS full-size view (for URIObject uploads)
  if (isAms) {
    urls.push(`${base}/views/imgpsh_fullsize_anim`);
  }
  // 2. Thumbnail/preview URL (AMS preview for SharePoint files, or URIObject thumbnail)
  if (attachment.thumbnailUrl) urls.push(attachment.thumbnailUrl);
  // 3. AMS base URL as last resort (for URIObject only — SharePoint URLs won't work with skype auth)
  if (isAms && !urls.includes(base)) urls.push(base);
  return urls;
}

function ImageAttachmentCard({
  attachment,
  tenantId,
}: {
  attachment: MessageAttachment;
  tenantId?: string | null;
}) {
  const candidateUrls = resolveImageFetchUrls(attachment);
  const [urlIndex, setUrlIndex] = useState(0);
  const containerRef = useRef<HTMLDivElement | null>(null);
  const [shouldLoad, setShouldLoad] = useState(false);
  const currentUrl = candidateUrls[urlIndex];
  const { src, loading } = useAuthenticatedImage(
    shouldLoad ? currentUrl : undefined,
    shouldLoad ? tenantId : undefined,
  );
  const [imgError, setImgError] = useState(false);
  const sizeLabel = formatAttachmentSize(attachment.fileSize);

  useEffect(() => {
    const node = containerRef.current;
    if (!node || shouldLoad) return;
    if (typeof IntersectionObserver !== "function") {
      setShouldLoad(true);
      return;
    }
    const observer = new IntersectionObserver(
      ([entry]) => {
        if (!entry?.isIntersecting) return;
        setShouldLoad(true);
        observer.disconnect();
      },
      { rootMargin: "300px 0px", threshold: 0.01 },
    );
    observer.observe(node);
    return () => observer.disconnect();
  }, [shouldLoad]);

  // If the <img> tag itself fails to load, try the next URL candidate
  const handleImgError = useCallback(() => {
    if (urlIndex < candidateUrls.length - 1) {
      setUrlIndex((i) => i + 1);
      setImgError(false);
    } else {
      setImgError(true);
    }
  }, [urlIndex, candidateUrls.length]);

  return (
    <div
      ref={containerRef}
      className="group/img relative mb-2 inline-block max-w-md overflow-hidden rounded-xl border border-border bg-muted/30"
    >
      <div className="flex h-72 w-full items-center justify-center bg-muted/40">
        {shouldLoad && loading ? (
          <span className="text-[12px] text-muted-foreground/50">Loading…</span>
        ) : !shouldLoad ? (
          <span className="text-[12px] text-muted-foreground/50">
            Loading preview…
          </span>
        ) : src && !imgError ? (
          <img
            src={src}
            alt={attachment.title}
            onError={handleImgError}
            className="h-full w-full bg-muted object-contain"
          />
        ) : (
          <span className="text-[13px] text-muted-foreground/50">
            Image unavailable
          </span>
        )}
      </div>
      <div className="pointer-events-none absolute top-2 right-2 flex gap-1.5 opacity-0 transition-opacity group-hover/img:pointer-events-auto group-hover/img:opacity-100">
        <button
          type="button"
          aria-label="Open in browser"
          onClick={() => openExternal(attachment.openUrl)}
          className="rounded-md border border-border/70 bg-background p-1.5 text-foreground/70 transition-colors hover:bg-accent hover:text-foreground"
        >
          <svg
            width="16"
            height="16"
            viewBox="0 0 16 16"
            fill="none"
            aria-hidden="true"
          >
            <path
              d="M6 3H3v10h10v-3M9 3h4v4M14 2L7 9"
              stroke="currentColor"
              strokeWidth="1.5"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          </svg>
        </button>
        <button
          type="button"
          aria-label="Download"
          onClick={() => openExternal(attachment.objectUrl)}
          className="rounded-md border border-border/70 bg-background p-1.5 text-foreground/70 transition-colors hover:bg-accent hover:text-foreground"
        >
          <DownloadIcon />
        </button>
      </div>

      {(attachment.title || sizeLabel) && (
        <div className="flex items-center justify-between gap-3 px-3 py-1.5">
          <span className="min-w-0 flex-1">
            <span className="block truncate text-[12px] text-muted-foreground/70">
              {attachment.title}
            </span>
          </span>
          {sizeLabel ? (
            <span className="shrink-0 text-[11px] text-muted-foreground/50">
              {sizeLabel}
            </span>
          ) : null}
        </div>
      )}
    </div>
  );
}

function FileAttachmentCard({ attachment }: { attachment: MessageAttachment }) {
  const sizeLabel = formatAttachmentSize(attachment.fileSize);
  const handleClick = useOpenExternal();

  return (
    <a
      href={attachment.openUrl}
      onClick={handleClick}
      className="mb-2 flex cursor-pointer items-center gap-3 rounded-xl border border-border bg-accent/30 px-3 py-2.5 transition-colors hover:bg-accent/50"
    >
      <span
        className={cn(
          "flex size-10 shrink-0 items-center justify-center rounded-lg text-[10px] font-bold tracking-wide",
          fileExtColor(attachment),
        )}
      >
        {fileExtLabel(attachment)}
      </span>
      <span className="min-w-0 flex-1">
        <span className="block truncate text-[13px] font-medium text-foreground">
          {attachment.title}
        </span>
        {sizeLabel ? (
          <span className="block text-[12px] text-muted-foreground/60">
            {sizeLabel}
          </span>
        ) : null}
      </span>
      <span className="flex shrink-0 items-center gap-2 text-[12px] text-muted-foreground">
        <DownloadIcon className="size-3.5" />
        Open
      </span>
    </a>
  );
}

function AttachmentCard({
  attachment,
  tenantId,
}: {
  attachment: MessageAttachment;
  tenantId?: string | null;
}) {
  if (attachment.kind === "image") {
    return <ImageAttachmentCard attachment={attachment} tenantId={tenantId} />;
  }
  return <FileAttachmentCard attachment={attachment} />;
}

// ── Inline Parts Utilities ──

function splitInlinePartsAtNewline(parts: MessageInlinePart[]): {
  before: MessageInlinePart[];
  after: MessageInlinePart[];
} | null {
  let remaining = parts;
  const before: MessageInlinePart[] = [];
  while (remaining.length > 0) {
    const [part, ...rest] = remaining;
    const newlineIdx = part.text.indexOf("\n");
    if (newlineIdx === -1) {
      before.push(part);
      remaining = rest;
      continue;
    }
    const beforePart = part.text.slice(0, newlineIdx);
    const afterPart = part.text.slice(newlineIdx + 1);
    if (beforePart) before.push({ ...part, text: beforePart });
    const after = [
      ...(afterPart ? [{ ...part, text: afterPart }] : []),
      ...rest,
    ] as MessageInlinePart[];
    return { before, after };
  }
  return null;
}

function inlineText(parts: MessageInlinePart[] | null | undefined): string {
  return parts?.map((part) => part.text).join("") ?? "";
}

function trimLeadingBlankLines(
  parts: MessageInlinePart[],
): MessageInlinePart[] {
  const next = [...parts];
  while (next.length > 0) {
    const first = next[0];
    if (!first) break;
    const trimmed = first.text.replace(/^\n+/, "");
    if (trimmed === first.text) break;
    if (trimmed.length === 0) {
      next.shift();
      continue;
    }
    next[0] = { ...first, text: trimmed };
    break;
  }
  return next;
}

// ── Rich Text Renderer ──

// ── Code Block ──

function CodeBlock({ text, language }: { text: string; language?: string }) {
  return (
    <div className="my-1.5 overflow-hidden rounded-lg border border-border bg-muted/60">
      {language ? (
        <div className="border-b border-border bg-muted/80 px-3 py-1 text-[11px] font-medium text-muted-foreground">
          {language}
        </div>
      ) : null}
      <pre className="overflow-x-auto p-3 text-[13px] leading-relaxed">
        <code className="font-mono text-foreground/90">{text}</code>
      </pre>
    </div>
  );
}

function RichText({
  parts,
  onOpenMessageRef,
  getMentionProfile,
  onOpenProfile,
}: {
  parts: MessageInlinePart[];
  onOpenMessageRef?: (conversationId: string, messageId: string) => void;
  getMentionProfile?: (part: MessageInlinePart) => ProfileData | null;
  onOpenProfile?: (profile: ProfileData) => void;
}) {
  // Split parts into segments: runs of inline parts vs code blocks
  const segments: Array<
    | { key: string; type: "inline"; parts: MessageInlinePart[] }
    | { key: string; type: "code_block"; text: string; language?: string }
  > = [];
  let currentInline: MessageInlinePart[] = [];
  let inlineStartOffset = 0;
  let cursor = 0;

  for (const part of parts) {
    if (part.kind === "code_block") {
      if (currentInline.length > 0) {
        segments.push({
          key: `inline-${inlineStartOffset}-${cursor}`,
          type: "inline",
          parts: currentInline,
        });
        currentInline = [];
      }
      segments.push({
        key: `code-${cursor}-${part.language ?? "plain"}-${part.text.length}`,
        type: "code_block",
        text: part.text,
        language: part.language,
      });
      cursor += part.text.length;
      inlineStartOffset = cursor;
    } else {
      if (currentInline.length === 0) inlineStartOffset = cursor;
      currentInline.push(part);
      cursor += part.text.length;
    }
  }
  if (currentInline.length > 0) {
    segments.push({
      key: `inline-${inlineStartOffset}-${cursor}`,
      type: "inline",
      parts: currentInline,
    });
  }

  return (
    <>
      {segments.map((segment) => {
        if (segment.type === "code_block") {
          return (
            <CodeBlock
              key={segment.key}
              text={segment.text}
              language={segment.language}
            />
          );
        }
        return (
          <InlineRichText
            key={segment.key}
            parts={segment.parts}
            onOpenMessageRef={onOpenMessageRef}
            getMentionProfile={getMentionProfile}
            onOpenProfile={onOpenProfile}
          />
        );
      })}
    </>
  );
}

function InlineRichText({
  parts,
  onOpenMessageRef,
  getMentionProfile,
  onOpenProfile,
}: {
  parts: MessageInlinePart[];
  onOpenMessageRef?: (conversationId: string, messageId: string) => void;
  getMentionProfile?: (part: MessageInlinePart) => ProfileData | null;
  onOpenProfile?: (profile: ProfileData) => void;
}) {
  const handleExternalClick = useOpenExternal();
  const lines: Array<{
    key: string;
    parts: Array<MessageInlinePart & { key: string }>;
  }> = [];
  let current: Array<MessageInlinePart & { key: string }> = [];
  let cursor = 0;
  let lineStart = 0;

  for (const part of parts) {
    const chunks = part.text.split("\n");
    for (let i = 0; i < chunks.length; i++) {
      const text = chunks[i];
      if (text) {
        current.push({
          ...part,
          text,
          key: `${cursor}-${part.kind}-${text}`,
        });
        cursor += text.length;
      }
      if (i < chunks.length - 1) {
        lines.push({ key: `line-${lineStart}`, parts: current });
        current = [];
        cursor += 1;
        lineStart = cursor;
      }
    }
  }
  lines.push({ key: `line-${lineStart}`, parts: current });

  while (lines.length > 1 && lines[0]?.parts.length === 0) {
    lines.shift();
  }
  while (lines.length > 1 && lines[lines.length - 1]?.parts.length === 0) {
    lines.pop();
  }

  return (
    <>
      {lines.map((line, lineIdx) => (
        <React.Fragment key={line.key}>
          {line.parts.map((part) => {
            if (part.kind === "code_block") return null; // handled by RichText
            const isCode = part.kind === "text" && part.code;
            const inlineClassName = cn(
              part.bold && "font-semibold",
              part.italic && "italic",
              part.strike && "line-through",
            );
            if (part.kind === "link") {
              return (
                <a
                  key={part.key}
                  href={part.href}
                  onClick={handleExternalClick}
                  className={cn(
                    "cursor-pointer font-medium text-primary underline decoration-primary/30 underline-offset-2 transition-colors hover:decoration-primary/60",
                    inlineClassName,
                  )}
                >
                  {part.text}
                </a>
              );
            }
            if (part.kind === "mention") {
              const mentionClassName =
                "rounded-[4px] bg-primary/15 px-1 py-0.5 text-[0.92em] font-semibold text-primary";
              if (part.messageRef) {
                const { conversationId, messageId } = part.messageRef;
                return (
                  <button
                    key={part.key}
                    type="button"
                    onClick={() =>
                      onOpenMessageRef?.(conversationId, messageId)
                    }
                    className={`${mentionClassName} cursor-pointer transition-all hover:bg-primary/20 focus-visible:bg-primary/20 focus-visible:outline-none`}
                  >
                    {part.text}
                  </button>
                );
              }
              if (part.href) {
                return (
                  <a
                    key={part.key}
                    href={part.href}
                    onClick={handleExternalClick}
                    className={mentionClassName}
                  >
                    {part.text}
                  </a>
                );
              }
              const mention = (
                <span
                  key={part.key}
                  className={cn(mentionClassName, inlineClassName)}
                >
                  {part.text}
                </span>
              );
              const profile = getMentionProfile?.(part) ?? null;
              return profile ? (
                <ProfileTrigger
                  key={part.key}
                  profile={profile}
                  onOpenProfile={onOpenProfile}
                >
                  {mention}
                </ProfileTrigger>
              ) : (
                mention
              );
            }
            if (isCode) {
              return (
                <code
                  key={part.key}
                  className={cn(
                    "rounded-[4px] bg-muted px-1.5 py-0.5 font-mono text-[0.88em] text-foreground/85",
                    inlineClassName,
                  )}
                >
                  {part.text}
                </code>
              );
            }
            return (
              <span key={part.key} className={inlineClassName}>
                {part.text}
              </span>
            );
          })}
          {lineIdx < lines.length - 1 ? <br /> : null}
        </React.Fragment>
      ))}
    </>
  );
}

// ── Quote Block ──

function QuoteBlock({
  parts,
  quoteRef,
  onOpenMessageRef,
  getMentionProfile,
  onOpenProfile,
}: {
  parts: MessageInlinePart[];
  quoteRef?: MessageReference | null;
  onOpenMessageRef?: (conversationId: string, messageId: string) => void;
  getMentionProfile?: (part: MessageInlinePart) => ProfileData | null;
  onOpenProfile?: (profile: ProfileData) => void;
}) {
  const split = splitInlinePartsAtNewline(parts);
  const firstLine = split ? inlineText(split.before) : inlineText(parts);
  const hasAuthor =
    split != null && firstLine.length > 0 && firstLine.length < 60;
  const author = hasAuthor ? (split?.before ?? null) : null;
  const body = author ? trimLeadingBlankLines(split?.after ?? []) : parts;

  const content = (
    <div className="mb-1.5 rounded-md border-l-[3px] border-muted-foreground/25 bg-accent/60 py-1.5 pr-3 pl-3 text-[13px] leading-relaxed text-muted-foreground">
      {author ? (
        <p className="mb-0.5 text-[12px] font-semibold text-foreground/80">
          <RichText
            parts={author}
            onOpenMessageRef={onOpenMessageRef}
            getMentionProfile={getMentionProfile}
            onOpenProfile={onOpenProfile}
          />
        </p>
      ) : null}
      <p className="break-words">
        <RichText
          parts={body}
          onOpenMessageRef={onOpenMessageRef}
          getMentionProfile={getMentionProfile}
          onOpenProfile={onOpenProfile}
        />
      </p>
    </div>
  );

  if (!quoteRef) return content;

  return (
    <button
      type="button"
      onClick={() =>
        onOpenMessageRef?.(quoteRef.conversationId, quoteRef.messageId)
      }
      className="block w-full cursor-pointer rounded-md text-left transition-all hover:bg-accent/40 focus-visible:bg-accent/40 focus-visible:outline-none"
    >
      {content}
    </button>
  );
}

// ── Deleted Message Placeholder ──

function DeletedMessageBody() {
  return (
    <p className="flex items-center gap-1.5 text-[13px] italic text-muted-foreground/50">
      <svg
        width="14"
        height="14"
        viewBox="0 0 16 16"
        fill="none"
        className="shrink-0"
        aria-hidden="true"
      >
        <circle cx="8" cy="8" r="6.5" stroke="currentColor" strokeWidth="1" />
        <path
          d="M5.5 5.5l5 5M10.5 5.5l-5 5"
          stroke="currentColor"
          strokeWidth="1"
          strokeLinecap="round"
        />
      </svg>
      This message has been deleted.
    </p>
  );
}

function sameSharedConversations(
  left: ProfileData["sharedConversations"],
  right: ProfileData["sharedConversations"],
): boolean {
  if (left === right) return true;
  if (!left || !right) return left === right;
  if (left.length !== right.length) return false;
  for (let i = 0; i < left.length; i++) {
    const leftConversation = left[i];
    const rightConversation = right[i];
    if (!leftConversation || !rightConversation) return false;
    if (
      leftConversation.id !== rightConversation.id ||
      leftConversation.kind !== rightConversation.kind ||
      leftConversation.title !== rightConversation.title
    ) {
      return false;
    }
  }
  return true;
}

function sameProfile(
  left?: ProfileData | null,
  right?: ProfileData | null,
): boolean {
  if (left === right) return true;
  if (!left || !right) return left === right;
  return (
    left.mri === right.mri &&
    left.displayName === right.displayName &&
    left.avatarThumbSrc === right.avatarThumbSrc &&
    left.avatarFullSrc === right.avatarFullSrc &&
    left.email === right.email &&
    left.jobTitle === right.jobTitle &&
    left.department === right.department &&
    left.companyName === right.companyName &&
    left.tenantName === right.tenantName &&
    left.location === right.location &&
    left.presence === right.presence &&
    left.currentConversationId === right.currentConversationId &&
    left.sharedConversationHeading === right.sharedConversationHeading &&
    sameSharedConversations(left.sharedConversations, right.sharedConversations)
  );
}

export function messageRowPropsEqual(
  prev: Readonly<React.ComponentProps<typeof MessageRowComponent>>,
  next: Readonly<React.ComponentProps<typeof MessageRowComponent>>,
): boolean {
  return (
    prev.entry === next.entry &&
    prev.avatarSrc === next.avatarSrc &&
    prev.showMeta === next.showMeta &&
    prev.presence === next.presence &&
    prev.isHighlighted === next.isHighlighted &&
    prev.tenantId === next.tenantId &&
    prev.onOpenMessageRef === next.onOpenMessageRef &&
    prev.onDeleteMessage === next.onDeleteMessage &&
    prev.getMentionProfile === next.getMentionProfile &&
    prev.onOpenProfile === next.onOpenProfile &&
    sameProfile(prev.profile, next.profile)
  );
}

function MessageRowComponent({
  entry,
  avatarSrc,
  showMeta,
  profile,
  presence,
  isHighlighted,
  tenantId,
  onOpenMessageRef,
  onDeleteMessage,
  getMentionProfile,
  onOpenProfile,
}: {
  entry: DisplayMessage;
  avatarSrc?: string;
  showMeta: boolean;
  profile?: ProfileData | null;
  presence?: PresenceInfo | null;
  isHighlighted?: boolean;
  tenantId?: string | null;
  onOpenMessageRef?: (conversationId: string, messageId: string) => void;
  onDeleteMessage?: (conversationId: string, messageId: string) => void;
  getMentionProfile?: (part: MessageInlinePart) => ProfileData | null;
  onOpenProfile?: (profile: ProfileData) => void;
}) {
  return (
    <li
      className="group/msg list-none"
      data-message-id={entry.message.id}
      data-highlighted={isHighlighted ? "true" : "false"}
    >
      <div
        className={cn(
          "relative flex gap-3 px-5 transition-colors duration-150",
          "group-hover/msg:bg-accent/40",
          showMeta ? "py-1 mt-2 pt-1" : "py-0.5",
          isHighlighted && "message-highlight-enter",
        )}
      >
        {/* Avatar column */}
        <div className={cn("w-14 shrink-0", showMeta ? "pt-0.5" : "pt-1")}>
          {showMeta ? (
            <ProfileTrigger
              profile={profile ?? null}
              onOpenProfile={onOpenProfile}
            >
              <MsgAvatar
                src={avatarSrc}
                name={entry.displayName}
                presence={entry.self ? null : presence}
              />
            </ProfileTrigger>
          ) : (
            <span className="block whitespace-nowrap text-[10px] leading-5 tabular-nums text-muted-foreground/0 group-hover/msg:text-muted-foreground/40">
              {entry.time}
            </span>
          )}
        </div>

        {/* Content column */}
        <div className="min-w-0 flex-1">
          {showMeta ? (
            <div className="flex items-baseline gap-2 pb-0.5">
              <ProfileTrigger
                profile={profile ?? null}
                onOpenProfile={onOpenProfile}
              >
                <span
                  className={cn(
                    "text-[15px] font-bold leading-snug tracking-[-0.01em] transition-colors hover:underline",
                    entry.self ? "text-primary" : "text-foreground",
                  )}
                >
                  {entry.displayName}
                </span>
              </ProfileTrigger>
              {entry.time ? (
                <span className="text-[12px] tabular-nums text-muted-foreground/50">
                  {entry.time}
                </span>
              ) : null}
              {entry.edited ? (
                <span className="text-[11px] text-muted-foreground/40">
                  (edited)
                </span>
              ) : null}
              {entry.self ? (
                <ReadReceiptDetails
                  status={entry.readStatus}
                  sentAt={entry.sentAt}
                  readAt={entry.readAt}
                  receiptScope={entry.receiptScope}
                  receiptSeenBy={entry.receiptSeenBy}
                  receiptUnseenBy={entry.receiptUnseenBy}
                />
              ) : null}
            </div>
          ) : null}

          {entry.deleted ? (
            <DeletedMessageBody />
          ) : (
            <div className="text-[15px] leading-[1.65] text-foreground/90">
              {entry.parts.quote ? (
                <QuoteBlock
                  parts={entry.parts.quote}
                  quoteRef={entry.parts.quoteRef}
                  onOpenMessageRef={onOpenMessageRef}
                  getMentionProfile={getMentionProfile}
                  onOpenProfile={onOpenProfile}
                />
              ) : null}
              {entry.parts.attachments.map((att) => (
                <AttachmentCard
                  key={
                    att.objectUrl || att.openUrl || att.fileName || att.title
                  }
                  attachment={att}
                  tenantId={tenantId}
                />
              ))}
              {entry.parts.body.length > 0 ? (
                <div className="break-words">
                  <RichText
                    parts={entry.parts.body}
                    onOpenMessageRef={onOpenMessageRef}
                    getMentionProfile={getMentionProfile}
                    onOpenProfile={onOpenProfile}
                  />
                </div>
              ) : null}
            </div>
          )}
        </div>

        {/* Hover actions */}
        {entry.self && !entry.deleted && onDeleteMessage ? (
          <div className="pointer-events-none absolute top-0 right-3 flex items-center gap-1 opacity-0 transition-opacity group-hover/msg:pointer-events-auto group-hover/msg:opacity-100">
            <button
              type="button"
              aria-label="Delete message"
              onClick={() =>
                onDeleteMessage(entry.message.conversationId, entry.message.id)
              }
              className="rounded-md p-1.5 text-muted-foreground/50 transition-colors hover:bg-destructive/10 hover:text-destructive"
            >
              <svg
                width="14"
                height="14"
                viewBox="0 0 16 16"
                fill="none"
                aria-hidden="true"
              >
                <path
                  d="M5 2h6M2 4h12M3.5 4l.7 9.1a1.5 1.5 0 0 0 1.5 1.4h4.6a1.5 1.5 0 0 0 1.5-1.4L12.5 4"
                  stroke="currentColor"
                  strokeWidth="1.2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            </button>
          </div>
        ) : null}
      </div>
    </li>
  );
}

export const MessageRow = React.memo(MessageRowComponent, messageRowPropsEqual);
