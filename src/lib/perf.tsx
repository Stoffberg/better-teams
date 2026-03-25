import { Profiler, type ReactNode } from "react";

export const PERF_FLAG = "better-teams-perf";

type PerfDetailValue = boolean | number | string | null;

export type PerfDetail = Record<string, PerfDetailValue>;

export type PerfMetric = {
  name: string;
  at: number;
  durationMs?: number;
  detail?: PerfDetail;
};

export type PerfSnapshot = {
  name: string;
  at: number;
  values: PerfDetail;
};

type PerfStore = {
  metrics: PerfMetric[];
  snapshots: Record<string, PerfSnapshot>;
};

declare global {
  interface Window {
    __BETTER_TEAMS_PERF__?: PerfStore;
  }
}

const PERF_METRIC_LIMIT = 500;

function now(): number {
  return typeof performance !== "undefined" ? performance.now() : Date.now();
}

export function isPerfEnabled(): boolean {
  if (typeof window === "undefined") return false;
  try {
    return window.localStorage.getItem(PERF_FLAG) === "1";
  } catch {
    return false;
  }
}

function getPerfStore(): PerfStore | null {
  if (typeof window === "undefined" || !isPerfEnabled()) return null;
  window.__BETTER_TEAMS_PERF__ ??= {
    metrics: [],
    snapshots: {},
  };
  return window.__BETTER_TEAMS_PERF__;
}

function pushMetric(metric: PerfMetric): void {
  const store = getPerfStore();
  if (!store) return;
  store.metrics.push(metric);
  if (store.metrics.length > PERF_METRIC_LIMIT) {
    store.metrics.splice(0, store.metrics.length - PERF_METRIC_LIMIT);
  }
}

export function recordPerfMetric(name: string, detail?: PerfDetail): void {
  pushMetric({
    name,
    at: now(),
    detail,
  });
}

export function updatePerfSnapshot(name: string, values: PerfDetail): void {
  const store = getPerfStore();
  if (!store) return;
  store.snapshots[name] = {
    name,
    at: now(),
    values,
  };
}

export function beginPerfMeasure(name: string, detail?: PerfDetail) {
  if (!isPerfEnabled()) {
    return (_endDetail?: PerfDetail) => undefined;
  }
  const startAt = now();
  pushMetric({
    name: `${name}.start`,
    at: startAt,
    detail,
  });
  return (endDetail?: PerfDetail) => {
    const endAt = now();
    pushMetric({
      name,
      at: endAt,
      durationMs: endAt - startAt,
      detail: {
        ...(detail ?? {}),
        ...(endDetail ?? {}),
      },
    });
  };
}

export async function measurePerfAsync<T>(
  name: string,
  detail: PerfDetail | undefined,
  run: () => Promise<T>,
): Promise<T> {
  const end = beginPerfMeasure(name, detail);
  try {
    const value = await run();
    end({ status: "ok" });
    return value;
  } catch (error) {
    end({
      status: "error",
      error:
        error instanceof Error
          ? error.name
          : typeof error === "string"
            ? error
            : "unknown",
    });
    throw error;
  }
}

export function countDomNodes(root: ParentNode | null | undefined): number {
  if (!root) return 0;
  if (root instanceof Element) {
    return root.querySelectorAll("*").length + 1;
  }
  return root.querySelectorAll("*").length;
}

export function resetPerfStore(): void {
  if (typeof window === "undefined") return;
  delete window.__BETTER_TEAMS_PERF__;
}

export function PerfProfiler({
  children,
  detail,
  id,
}: {
  children: ReactNode;
  detail?: PerfDetail;
  id: string;
}) {
  if (!isPerfEnabled()) {
    return <>{children}</>;
  }
  return (
    <Profiler
      id={id}
      onRender={(
        profilerId,
        phase,
        actualDuration,
        baseDuration,
        startTime,
        commitTime,
      ) => {
        pushMetric({
          name: "react.profiler",
          at: commitTime,
          durationMs: actualDuration,
          detail: {
            id: profilerId,
            phase,
            actualDurationMs: Number(actualDuration.toFixed(3)),
            baseDurationMs: Number(baseDuration.toFixed(3)),
            startTimeMs: Number(startTime.toFixed(3)),
            commitTimeMs: Number(commitTime.toFixed(3)),
            ...(detail ?? {}),
          },
        });
      }}
    >
      {children}
    </Profiler>
  );
}
