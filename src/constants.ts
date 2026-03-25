export const APP_NAME = "Better Teams";
export const APP_SLUG = "better-teams";
export const APP_DESCRIPTION =
  "A calmer, faster Microsoft Teams desktop experience with a focused custom UI.";

export const DEV_RENDERER_PORT = 5173;

export const DEV_RENDERER_URL = `http://localhost:${DEV_RENDERER_PORT}`;

export function isAllowedNavigationHost(hostname: string): boolean {
  return hostname === "localhost" || hostname === "127.0.0.1";
}

export const ALLOWED_PERMISSIONS = [
  "media",
  "notifications",
  "clipboard-read",
  "clipboard-sanitized-write",
  "screen",
];
