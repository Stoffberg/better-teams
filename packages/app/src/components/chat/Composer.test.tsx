import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { fireEvent, render, screen } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import { Composer } from "./Composer";

vi.mock("@better-teams/core/teams-client-factory");

function renderComposer() {
  const queryClient = new QueryClient({
    defaultOptions: {
      queries: {
        retry: false,
      },
    },
  });

  return render(
    <QueryClientProvider client={queryClient}>
      <Composer
        conversationId="c1"
        conversationTitle="Design chat"
        conversationMembers={[]}
        composerRef={{ current: document.createElement("div") }}
        liveSessionReady
        mentionCandidates={[]}
      />
    </QueryClientProvider>,
  );
}

describe("Composer", () => {
  it("keeps a single link control in the bottom action bar", () => {
    renderComposer();

    expect(screen.getAllByRole("button", { name: /add link/i })).toHaveLength(
      1,
    );
    expect(
      screen.queryByRole("button", { name: /^video$/i }),
    ).not.toBeInTheDocument();
    expect(
      screen.queryByRole("button", { name: /^audio$/i }),
    ).not.toBeInTheDocument();
    expect(
      screen.queryByRole("button", { name: /^checklist$/i }),
    ).not.toBeInTheDocument();
  });

  it("styles lists in the editor so bullets and numbers render", () => {
    renderComposer();

    expect(screen.getByLabelText(/message text/i).className).toContain(
      "[&_ul]:list-disc",
    );
    expect(screen.getByLabelText(/message text/i).className).toContain(
      "[&_ol]:list-decimal",
    );
  });

  it("runs the bullet list command from the toolbar", () => {
    const execCommand = vi.fn(() => true);
    document.execCommand = execCommand;

    renderComposer();
    fireEvent.click(screen.getByRole("button", { name: /bullet list/i }));

    expect(execCommand).toHaveBeenCalledWith(
      "insertUnorderedList",
      false,
      undefined,
    );
  });
});
