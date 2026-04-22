import { describe, expect, it } from "vitest";
import {
  FetchRequestSchema,
  FetchResponseSchema,
  ImageCacheIpcRequestSchema,
  PresenceRequestSchema,
  RawTokenSchema,
  ShellOpenExternalUrlSchema,
} from "./contracts";

describe("desktop preload contracts", () => {
  it("accepts valid token payloads", () => {
    expect(
      RawTokenSchema.parse({
        host: "teams.microsoft.com",
        name: "skypetoken",
        token: "token",
        expiresAt: "2026-03-25T00:00:00Z",
      }),
    ).toMatchObject({ token: "token" });
  });

  it("fails closed on malformed token payloads", () => {
    expect(() =>
      RawTokenSchema.parse({
        host: "teams.microsoft.com",
        name: "skypetoken",
        expiresAt: "2026-03-25T00:00:00Z",
      }),
    ).toThrow();
  });

  it("rejects invalid presence request payloads", () => {
    expect(() => PresenceRequestSchema.parse(["8:orgid:one"])).not.toThrow();
    expect(() => PresenceRequestSchema.parse(["8:orgid:one", 42])).toThrow();
  });

  it("rejects invalid image cache IPC payloads", () => {
    expect(() =>
      ImageCacheIpcRequestSchema.parse({
        cacheKey: "avatar",
        bytes: [0, 255],
        extension: "png",
      }),
    ).not.toThrow();
    expect(() =>
      ImageCacheIpcRequestSchema.parse({
        cacheKey: "avatar",
        bytes: [256],
        extension: "png",
      }),
    ).toThrow();
  });

  it("rejects invalid fetch bridge payloads", () => {
    expect(() =>
      FetchRequestSchema.parse({
        url: "https://teams.microsoft.com/api",
        headers: [["accept", "application/json"]],
        body: null,
      }),
    ).not.toThrow();
    expect(() =>
      FetchRequestSchema.parse({
        url: "not a url",
        headers: [["accept", "application/json"]],
        body: null,
      }),
    ).toThrow();
    expect(() =>
      FetchResponseSchema.parse({
        status: 200,
        statusText: "OK",
        headers: [["content-type", "application/json"]],
        body: "not bytes",
      }),
    ).toThrow();
  });

  it("rejects invalid shell open targets", () => {
    expect(() =>
      ShellOpenExternalUrlSchema.parse("https://teams.microsoft.com"),
    ).not.toThrow();
    expect(() =>
      ShellOpenExternalUrlSchema.parse("javascript:alert(1)"),
    ).toThrow();
  });
});
