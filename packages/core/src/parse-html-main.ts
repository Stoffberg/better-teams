import { parseHTML } from "linkedom";

export function parseHtmlDocument(html: string): Document {
  const { document } = parseHTML(
    `<!DOCTYPE html><html><head></head><body>${html}</body></html>`,
  );
  return document as unknown as Document;
}
