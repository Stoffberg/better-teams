import { parseHtmlDocument as parseHtmlInBrowser } from "./parse-html-browser";
import { parseHtmlDocument as parseHtmlInMain } from "./parse-html-main";

export function parseHtmlDocument(html: string): Document {
  const viteMain =
    typeof import.meta !== "undefined" &&
    import.meta.env &&
    Boolean(import.meta.env.VITE_ELECTRON_MAIN);
  if (viteMain) {
    return parseHtmlInMain(html);
  }
  if (typeof globalThis.DOMParser === "undefined") {
    return parseHtmlInMain(html);
  }
  return parseHtmlInBrowser(html);
}
