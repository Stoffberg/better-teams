import { describe, expect, it } from "vitest";
import { DEV_RENDERER_URL, isAllowedNavigationHost } from "./constants";

describe("constants", () => {
  it("dev renderer url points at local vite port", () => {
    expect(DEV_RENDERER_URL).toBe("http://localhost:5173");
  });

  it("allows only local dev hosts for navigation", () => {
    expect(isAllowedNavigationHost("localhost")).toBe(true);
    expect(isAllowedNavigationHost("127.0.0.1")).toBe(true);
    expect(isAllowedNavigationHost("teams.microsoft.com")).toBe(false);
    expect(isAllowedNavigationHost("evil.example")).toBe(false);
  });
});
