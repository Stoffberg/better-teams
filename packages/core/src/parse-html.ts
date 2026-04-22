import { parseHtmlDocument as parseHtmlInBrowser } from "./parse-html-browser";
import { parseHtmlDocument as parseHtmlInMain } from "./parse-html-main";

export function parseHtmlDocument(html: string): Document {
  const meta = import.meta as ImportMeta & {
    env?: { VITE_ELECTRON_MAIN?: string | boolean };
  };
  const viteMain =
    typeof import.meta !== "undefined" &&
    meta.env &&
    Boolean(meta.env.VITE_ELECTRON_MAIN);
  if (viteMain) {
    return parseHtmlInMain(html);
  }
  if (typeof globalThis.DOMParser === "undefined") {
    return parseHtmlInMain(html);
  }
  return parseHtmlInBrowser(html);
}
