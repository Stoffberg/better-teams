import { teamsKeys } from "@better-teams/app/lib/teams-query-keys";
import type { Message } from "@better-teams/core/teams/types";
import { getOrCreateClient } from "@better-teams/core/teams-client-factory";
import {
  sortMessagesByTimestamp,
  type ThreadQueryData,
} from "@better-teams/core/thread";
import type { QueryClient } from "@tanstack/react-query";
import {
  type RefObject,
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from "react";
import { OLDER_LOAD_THROTTLE_MS } from "./types";

const OLDER_PREFETCH_THRESHOLD_PX = 1800;
const OLDER_PREFETCH_ROOT_MARGIN = "2200px 0px 0px 0px";
const SCROLL_RESTORE_EPSILON_PX = 0.75;
const EXPECTED_OLDER_FETCH_MS = 1500;
const MAX_VELOCITY_PREFETCH_BONUS_PX = 6000;

type ScrollRestoreAnchor = {
  messageId: string | null;
  scrollHeight: number;
  scrollTop: number;
  top: number;
};

type UseThreadScrollOptions = {
  tenantId?: string | null;
  conversationId: string;
  queryClient: QueryClient;
  viewportRef: RefObject<HTMLDivElement | null>;
  topSentinelRef: RefObject<HTMLLIElement | null>;
  threadLoading: boolean;
  threadHasData: boolean;
  tailMessageId: string | null;
  loadedMessageCount: number;
  rawMessages: Message[];
  scrollToMessage: (messageId: string) => void;
};

function firstVisibleMessageNode(viewport: HTMLElement): HTMLElement | null {
  const viewportTop = viewport.getBoundingClientRect().top;
  const messageNodes =
    viewport.querySelectorAll<HTMLElement>("[data-message-id]");
  for (const node of messageNodes) {
    if (node.getBoundingClientRect().bottom > viewportTop) {
      return node;
    }
  }
  return null;
}

export function captureScrollRestoreAnchor(
  viewport: HTMLElement,
): ScrollRestoreAnchor {
  const anchorNode = firstVisibleMessageNode(viewport);
  const viewportTop = viewport.getBoundingClientRect().top;
  return {
    messageId: anchorNode?.dataset.messageId ?? null,
    scrollHeight: viewport.scrollHeight,
    scrollTop: viewport.scrollTop,
    top: anchorNode ? anchorNode.getBoundingClientRect().top - viewportTop : 0,
  };
}

export function restoreScrollRestoreAnchor(
  viewport: HTMLElement,
  restore: ScrollRestoreAnchor,
): void {
  if (restore.messageId) {
    const messageNodes =
      viewport.querySelectorAll<HTMLElement>("[data-message-id]");
    const anchorNode =
      [...messageNodes].find(
        (node) => node.dataset.messageId === restore.messageId,
      ) ?? null;
    if (anchorNode) {
      const viewportTop = viewport.getBoundingClientRect().top;
      const nextTop = anchorNode.getBoundingClientRect().top - viewportTop;
      const delta = nextTop - restore.top;
      if (Math.abs(delta) <= SCROLL_RESTORE_EPSILON_PX) return;
      viewport.scrollTop += delta;
      return;
    }
  }
  const targetScrollTop =
    viewport.scrollHeight - restore.scrollHeight + restore.scrollTop;
  if (
    Math.abs(targetScrollTop - viewport.scrollTop) <= SCROLL_RESTORE_EPSILON_PX
  ) {
    return;
  }
  viewport.scrollTop = targetScrollTop;
}

export function olderPrefetchThresholdForVelocity(
  upwardVelocityPxPerMs: number,
): number {
  const velocityBonus = Math.min(
    MAX_VELOCITY_PREFETCH_BONUS_PX,
    Math.max(0, upwardVelocityPxPerMs) * EXPECTED_OLDER_FETCH_MS,
  );
  return OLDER_PREFETCH_THRESHOLD_PX + velocityBonus;
}

export function shouldPrefetchOlderMessages(
  scrollTop: number,
  upwardVelocityPxPerMs = 0,
): boolean {
  return scrollTop <= olderPrefetchThresholdForVelocity(upwardVelocityPxPerMs);
}

export function useThreadScroll({
  tenantId,
  conversationId,
  queryClient,
  viewportRef,
  topSentinelRef,
  threadLoading,
  threadHasData,
  tailMessageId,
  loadedMessageCount,
  rawMessages,
  scrollToMessage,
}: UseThreadScrollOptions): {
  loadingOlder: boolean;
  onScroll: () => void;
  setPendingScrollMessageId: (messageId: string) => void;
} {
  const [loadingOlder, setLoadingOlder] = useState(false);
  const prevLastMessageIdRef = useRef<string | null>(null);
  const scrollRestoreRef = useRef<ScrollRestoreAnchor | null>(null);
  const lastOlderLoadAtRef = useRef(0);
  const lastScrollSampleRef = useRef<{ top: number; time: number } | null>(
    null,
  );
  const loadingOlderRef = useRef(false);
  const scrollFrameRef = useRef<number | null>(null);
  const pendingScrollMessageIdRef = useRef<string | null>(null);

  const loadOlderMessages = useCallback(async () => {
    if (loadingOlderRef.current) return;
    const snapshot = queryClient.getQueryData<ThreadQueryData>(
      teamsKeys.thread(tenantId, conversationId),
    );
    if (!snapshot?.moreOlder || !snapshot.olderPageUrl) return;
    const now = Date.now();
    if (now - lastOlderLoadAtRef.current < OLDER_LOAD_THROTTLE_MS) return;
    lastOlderLoadAtRef.current = now;
    loadingOlderRef.current = true;
    setLoadingOlder(true);
    try {
      const client = await getOrCreateClient(tenantId ?? undefined);
      const res = await client.getMessagesByUrl(snapshot.olderPageUrl);
      const batch = res.messages ?? [];
      const el = viewportRef.current;
      scrollRestoreRef.current = el ? captureScrollRestoreAnchor(el) : null;
      queryClient.setQueryData<ThreadQueryData>(
        teamsKeys.thread(tenantId, conversationId),
        (old) => {
          if (!old) return old;
          const merged = new Map(old.messages.map((m) => [m.id, m]));
          let added = 0;
          for (const m of batch) {
            if (!merged.has(m.id)) added++;
            merged.set(m.id, m);
          }
          if (added === 0) return { ...old, moreOlder: false };
          const messages = sortMessagesByTimestamp([...merged.values()]);
          const nextUrl = res._metadata?.backwardLink ?? null;
          return {
            messages,
            olderPageUrl: nextUrl,
            moreOlder: nextUrl != null,
          };
        },
      );
    } catch {
      scrollRestoreRef.current = null;
    } finally {
      loadingOlderRef.current = false;
      setLoadingOlder(false);
    }
  }, [conversationId, queryClient, tenantId, viewportRef]);

  useEffect(() => {
    const sentinel = topSentinelRef.current;
    const viewport = viewportRef.current;
    if (!sentinel || !viewport) return;
    const observer = new IntersectionObserver(
      ([entry]) => {
        if (entry?.isIntersecting) void loadOlderMessages();
      },
      {
        root: viewport,
        rootMargin: OLDER_PREFETCH_ROOT_MARGIN,
        threshold: 0,
      },
    );
    observer.observe(sentinel);
    return () => observer.disconnect();
  }, [loadOlderMessages, topSentinelRef, viewportRef]);

  const onScroll = useCallback(() => {
    const now = performance.now();
    const viewport = viewportRef.current;
    let upwardVelocityPxPerMs = 0;
    if (viewport) {
      const lastSample = lastScrollSampleRef.current;
      if (lastSample) {
        const deltaTime = Math.max(1, now - lastSample.time);
        const deltaTop = lastSample.top - viewport.scrollTop;
        upwardVelocityPxPerMs = deltaTop > 0 ? deltaTop / deltaTime : 0;
      }
      lastScrollSampleRef.current = {
        top: viewport.scrollTop,
        time: now,
      };
    }
    if (scrollFrameRef.current != null) return;
    scrollFrameRef.current = window.requestAnimationFrame(() => {
      scrollFrameRef.current = null;
      const el = viewportRef.current;
      if (!el) return;
      if (!shouldPrefetchOlderMessages(el.scrollTop, upwardVelocityPxPerMs)) {
        return;
      }
      void loadOlderMessages();
    });
  }, [loadOlderMessages, viewportRef]);

  useLayoutEffect(() => {
    if (threadLoading && !threadHasData) return;
    const el = viewportRef.current;
    if (!el) return;
    const last = tailMessageId;
    const prev = prevLastMessageIdRef.current;
    prevLastMessageIdRef.current = last;
    if (last === null) return;
    if (prev === null) {
      el.scrollTop = el.scrollHeight;
      return;
    }
    if (prev !== last) {
      const atBottom = el.scrollHeight - el.scrollTop - el.clientHeight < 150;
      if (atBottom) {
        el.scrollTop = el.scrollHeight;
      }
    }
  }, [threadHasData, threadLoading, tailMessageId, viewportRef]);

  useLayoutEffect(() => {
    const el = viewportRef.current;
    if (!el) return;
    if (loadedMessageCount === 0) return;
    const restore = scrollRestoreRef.current;
    if (!restore) return;
    scrollRestoreRef.current = null;
    restoreScrollRestoreAnchor(el, restore);
  }, [loadedMessageCount, viewportRef]);

  useEffect(() => {
    if (loadingOlderRef.current) return;
    const el = viewportRef.current;
    if (loadedMessageCount === 0) return;
    if (!el || !shouldPrefetchOlderMessages(el.scrollTop)) return;
    const snapshot = queryClient.getQueryData<ThreadQueryData>(
      teamsKeys.thread(tenantId, conversationId),
    );
    if (!snapshot?.moreOlder || !snapshot.olderPageUrl) return;
    const frame = window.requestAnimationFrame(() => {
      void loadOlderMessages();
    });
    return () => window.cancelAnimationFrame(frame);
  }, [
    conversationId,
    loadOlderMessages,
    loadedMessageCount,
    queryClient,
    tenantId,
    viewportRef,
  ]);

  useLayoutEffect(() => {
    const pendingId = pendingScrollMessageIdRef.current;
    if (!pendingId) return;
    if (!rawMessages.some((message) => message.id === pendingId)) return;
    pendingScrollMessageIdRef.current = null;
    scrollToMessage(pendingId);
  }, [rawMessages, scrollToMessage]);

  useEffect(() => {
    return () => {
      if (scrollFrameRef.current != null) {
        window.cancelAnimationFrame(scrollFrameRef.current);
      }
    };
  }, []);

  return {
    loadingOlder,
    onScroll,
    setPendingScrollMessageId(messageId: string) {
      pendingScrollMessageIdRef.current = messageId;
    },
  };
}
