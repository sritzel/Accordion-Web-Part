import * as React from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';

interface IErrorBoundaryProps {
  // Optional in the type so we can construct via React.createElement with
  // children passed as a positional arg. At runtime children is always provided.
  children?: React.ReactNode;
  // When in Edit mode the full error message is shown inline to help authors
  // diagnose. In Read mode (end users) we show only a generic message to
  // avoid leaking implementation details like list GUIDs or field names.
  displayMode: DisplayMode;
}

interface IErrorBoundaryState {
  error: Error | undefined;
}

/**
 * Small React error boundary. If anything below it throws during render, we
 * surface a friendly message inline. The full message (with stack details)
 * is only shown to page editors; end users see a generic notice. The full
 * error is always logged to the browser console for developer access.
 */
export class ErrorBoundary extends React.Component<IErrorBoundaryProps, IErrorBoundaryState> {
  public state: IErrorBoundaryState = { error: undefined };

  public static getDerivedStateFromError(error: Error): IErrorBoundaryState {
    return { error };
  }

  public componentDidCatch(error: Error, info: React.ErrorInfo): void {
    // Surface the details in the dev console for debugging.
    // eslint-disable-next-line no-console
    console.error('FaqAccordion error:', error, info);
  }

  public render(): React.ReactNode {
    if (this.state.error) {
      const isEdit = this.props.displayMode === DisplayMode.Edit;
      return (
        <div style={{
          padding: '12px 16px',
          border: '1px solid #a4262c',
          background: '#fde7e9',
          color: '#a4262c',
          fontFamily: '"Segoe UI", sans-serif',
          fontSize: '14px',
          borderRadius: '4px'
        }}>
          <div style={{ fontWeight: 600, marginBottom: '4px' }}>
            FAQ Accordion couldn&apos;t be displayed
          </div>
          {isEdit ? (
            <div style={{ fontFamily: 'monospace', whiteSpace: 'pre-wrap' }}>
              {this.state.error.message || String(this.state.error)}
            </div>
          ) : (
            <div>Please contact the site administrator if this persists.</div>
          )}
        </div>
      );
    }
    return this.props.children;
  }
}
