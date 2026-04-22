import { Button } from "@better-teams/ui/components/button";
import { Component, type ErrorInfo, type ReactNode } from "react";

type Props = { children: ReactNode };

type State = { error: Error | null };

export class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null };

  static getDerivedStateFromError(error: Error): State {
    return { error };
  }

  override componentDidCatch(error: Error, info: ErrorInfo): void {
    console.error(error, info.componentStack);
  }

  override render(): ReactNode {
    if (this.state.error) {
      return (
        <div className="space-y-4 p-6" role="alert">
          <h1 className="font-semibold text-2xl">Something broke</h1>
          <p className="text-muted-foreground text-sm">
            {this.state.error.message}
          </p>
          <Button
            type="button"
            onClick={() => {
              this.setState({ error: null });
            }}
          >
            Try again
          </Button>
        </div>
      );
    }
    return this.props.children;
  }
}
