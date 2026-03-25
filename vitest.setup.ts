import "@testing-library/jest-dom/vitest";
import { vi } from "vitest";

const localStorageMemory = new Map<string, string>();
Object.defineProperty(globalThis, "localStorage", {
  value: {
    get length() {
      return localStorageMemory.size;
    },
    clear() {
      localStorageMemory.clear();
    },
    getItem(key: string) {
      const v = localStorageMemory.get(key);
      return v === undefined ? null : v;
    },
    key(index: number) {
      return Array.from(localStorageMemory.keys())[index] ?? null;
    },
    removeItem(key: string) {
      localStorageMemory.delete(key);
    },
    setItem(key: string, value: string) {
      localStorageMemory.set(key, value);
    },
  },
  configurable: true,
});

// ── Mock Tauri modules ──
// These prevent runtime errors when tests import code that uses Tauri APIs.

vi.mock("@tauri-apps/api/core", () => ({
  invoke: vi.fn().mockResolvedValue(null),
}));

vi.mock("@tauri-apps/plugin-http", () => ({
  fetch: vi.fn().mockResolvedValue(new Response("{}", { status: 200 })),
}));

vi.mock("@tauri-apps/plugin-sql", () => {
  const mockDb = {
    execute: vi.fn().mockResolvedValue(undefined),
    select: vi.fn().mockResolvedValue([]),
    close: vi.fn().mockResolvedValue(undefined),
  };
  return {
    default: {
      load: vi.fn().mockResolvedValue(mockDb),
    },
  };
});
