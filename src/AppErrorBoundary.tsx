import { Component } from "react";
import type { ErrorInfo, ReactNode } from "react";

type Props = { children: ReactNode };
type State = { hasError: boolean };

export class AppErrorBoundary extends Component<Props, State> {
  state: State = { hasError: false };

  static getDerivedStateFromError(): State {
    return { hasError: true };
  }

  componentDidCatch(error: Error, _info: ErrorInfo) {
    // Do not log financial records, component props, stack traces, or user data.
    console.error("Budgetwise UI rendering failed", { errorType: error.name || "UnknownError" });
  }

  render() {
    if (this.state.hasError) {
      return (
        <main className="fatal-error" role="alert">
          <section className="fatal-error-card">
            <p className="overline">BUDGETWISE</p>
            <h1>Something went wrong</h1>
            <p>Your financial records have not been included in this error message. Reload the app to try again.</p>
            <button className="btn btn-primary" type="button" onClick={() => window.location.reload()}>
              Reload app
            </button>
          </section>
        </main>
      );
    }
    return this.props.children;
  }
}
