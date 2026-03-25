import { Component, type ErrorInfo, type ReactNode } from "react";

type Props = { label: string; children: ReactNode };
type State = { error: Error | null };

export class ChatErrorBoundary extends Component<Props, State> {
  state: State = { error: null };

  static getDerivedStateFromError(error: Error): State {
    return { error };
  }

  override componentDidCatch(error: Error, info: ErrorInfo): void {
    console.error(`[${this.props.label}]`, error, info.componentStack);
  }

  override render(): ReactNode {
    if (this.state.error) {
      return (
        <div
          className="flex flex-1 flex-col items-center justify-center gap-3 p-8"
          role="alert"
        >
          <p className="text-[13px] text-muted-foreground">
            Something went wrong in {this.props.label.toLowerCase()}.
          </p>
          <button
            type="button"
            onClick={() => this.setState({ error: null })}
            className="rounded-md border px-3 py-1.5 text-[12px] transition-colors hover:bg-accent"
          >
            Try again
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}
