import { fireEvent, render, screen } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import { ErrorBoundary } from "./ErrorBoundary";

function Boom(): null {
  throw new Error("unit test boom");
}

const flaky = { throwError: true };

function Flaky() {
  if (flaky.throwError) {
    throw new Error("bad");
  }
  return <span>recovered</span>;
}

describe("ErrorBoundary", () => {
  it("renders fallback when a child throws", () => {
    vi.spyOn(console, "error").mockImplementation(() => {});
    render(
      <ErrorBoundary>
        <Boom />
      </ErrorBoundary>,
    );
    expect(screen.getByRole("alert")).toBeInTheDocument();
    expect(screen.getByText(/unit test boom/)).toBeInTheDocument();
    vi.restoreAllMocks();
  });

  it("clears after Try again when the child no longer throws", () => {
    vi.spyOn(console, "error").mockImplementation(() => {});
    flaky.throwError = true;
    render(
      <ErrorBoundary>
        <Flaky />
      </ErrorBoundary>,
    );
    expect(screen.getByRole("alert")).toBeInTheDocument();
    flaky.throwError = false;
    fireEvent.click(screen.getByRole("button", { name: /try again/i }));
    expect(screen.getByText("recovered")).toBeInTheDocument();
    vi.restoreAllMocks();
  });
});
