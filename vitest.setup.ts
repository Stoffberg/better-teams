import "@testing-library/jest-dom/vitest";

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

Object.defineProperty(window, "betterTeams", {
  value: undefined,
  writable: true,
  configurable: true,
});
