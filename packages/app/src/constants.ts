const DEV_RENDERER_PORT = 5173;

export const DEV_RENDERER_URL = `http://localhost:${DEV_RENDERER_PORT}`;

export function isAllowedNavigationHost(hostname: string): boolean {
  return hostname === "localhost" || hostname === "127.0.0.1";
}
