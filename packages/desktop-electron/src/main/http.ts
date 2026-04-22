type FetchRequest = {
  url: string;
  method?: string;
  headers: [string, string][];
  body: string | ArrayBuffer | null;
};

type FetchResponse = {
  status: number;
  statusText: string;
  headers: [string, string][];
  body: ArrayBuffer;
};

export async function performFetch(
  request: FetchRequest,
): Promise<FetchResponse> {
  const response = await fetch(request.url, {
    method: request.method,
    headers: Object.fromEntries(request.headers),
    body: request.body ? request.body : undefined,
  });

  return {
    status: response.status,
    statusText: response.statusText,
    headers: [...response.headers.entries()],
    body: await response.arrayBuffer(),
  };
}
