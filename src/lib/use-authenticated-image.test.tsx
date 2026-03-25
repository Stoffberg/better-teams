import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { render, screen, waitFor } from "@testing-library/react";
import type { ReactNode } from "react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { useAuthenticatedImage } from "./use-authenticated-image";

vi.mock("./teams-client-factory", () => ({
  getOrCreateClient: vi.fn(),
}));

import { getOrCreateClient } from "./teams-client-factory";

function TestImage({
  imageUrl,
  tenantId,
}: {
  imageUrl?: string;
  tenantId?: string | null;
}) {
  const { src, loading, error } = useAuthenticatedImage(imageUrl, tenantId);
  return (
    <div>
      <span data-testid="src">{src ?? ""}</span>
      <span data-testid="loading">{String(loading)}</span>
      <span data-testid="error">{String(error)}</span>
    </div>
  );
}

function renderWithQueryClient(node: ReactNode) {
  const client = new QueryClient({
    defaultOptions: {
      queries: {
        retry: false,
        gcTime: 0,
        refetchOnWindowFocus: false,
      },
    },
  });

  return render(
    <QueryClientProvider client={client}>{node}</QueryClientProvider>,
  );
}

describe("useAuthenticatedImage", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("keeps the image source empty when authenticated fetch returns null", async () => {
    vi.mocked(getOrCreateClient).mockResolvedValue({
      fetchAuthenticatedImageSrc: vi.fn().mockResolvedValue(null),
    } as never);

    renderWithQueryClient(
      <TestImage
        imageUrl="https://backupdirect-my.sharepoint.com/image.png"
        tenantId="t1"
      />,
    );

    await waitFor(() => {
      expect(screen.getByTestId("loading")).toHaveTextContent("false");
    });

    expect(screen.getByTestId("src")).toHaveTextContent("");
    expect(screen.getByTestId("error")).toHaveTextContent("true");
  });

  it("returns the cached authenticated asset url when fetch succeeds", async () => {
    vi.mocked(getOrCreateClient).mockResolvedValue({
      fetchAuthenticatedImageSrc: vi
        .fn()
        .mockResolvedValue("asset:///tmp/image.png"),
    } as never);

    renderWithQueryClient(
      <TestImage
        imageUrl="https://eu-api.asm.skype.com/v1/objects/1/views/imgo"
        tenantId="t1"
      />,
    );

    await waitFor(() => {
      expect(screen.getByTestId("src")).toHaveTextContent(
        "asset:///tmp/image.png",
      );
    });

    expect(screen.getByTestId("loading")).toHaveTextContent("false");
    expect(screen.getByTestId("error")).toHaveTextContent("false");
  });
});
