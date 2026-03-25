import { openUrl } from "@tauri-apps/plugin-opener";

/**
 * Open a URL in the user's default browser via Tauri's opener plugin.
 * Falls back to window.open for dev/test environments where Tauri isn't available.
 */
export async function openExternal(url: string): Promise<void> {
  try {
    await openUrl(url);
  } catch {
    // Fallback for non-Tauri environments (dev server in browser, tests)
    window.open(url, "_blank", "noopener,noreferrer");
  }
}
