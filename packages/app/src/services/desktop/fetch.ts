import type { BetterTeamsDesktopApi } from "@better-teams/desktop-electron/preload";

type FetchInput = Parameters<typeof globalThis.fetch>[0];
type FetchInit = Parameters<typeof globalThis.fetch>[1];

export async function fetch(
  input: FetchInput,
  init?: FetchInit,
): Promise<Response> {
  const response = await api().http.fetch({
    url: resolveUrl(input),
    method: init?.method,
    headers: [...new Headers(init?.headers).entries()],
    body: await serializeBody(init?.body),
  });

  return new Response(response.body, {
    status: response.status,
    statusText: response.statusText,
    headers: response.headers,
  });
}

async function serializeBody(
  body: BodyInit | null | undefined,
): Promise<string | ArrayBuffer | null> {
  if (!body) return null;
  if (typeof body === "string") return body;
  if (body instanceof URLSearchParams) return body.toString();
  if (body instanceof ArrayBuffer) return body;
  if (ArrayBuffer.isView(body)) {
    return body.buffer.slice(
      body.byteOffset,
      body.byteOffset + body.byteLength,
    ) as ArrayBuffer;
  }
  if (body instanceof Blob) return body.arrayBuffer();
  throw new Error("Unsupported Electron fetch body type");
}

function resolveUrl(input: FetchInput): string {
  if (typeof input === "string") return input;
  if (input instanceof URL) return input.href;
  return input.url;
}

function api(): BetterTeamsDesktopApi {
  if (!window.betterTeams) {
    throw new Error("Better Teams desktop API is not available");
  }
  return window.betterTeams;
}
