import * as React from 'react';

interface IErrorBoundaryProps {
  children: React.ReactNode;
}

interface IErrorBoundaryState {
  error: Error | undefined;
}

/**
 * Small React error boundary. If anything below it throws during render, we
 * surface the real message inline instead of letting SPFx's generic card
 * render "ERROR: [object Object]" and hide the stack.
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
          <div style={{ fontWeight: 600, marginBottom: '4px' }}>FAQ Accordion failed to render</div>
          <div style={{ fontFamily: 'monospace', whiteSpace: 'pre-wrap' }}>
            {this.state.error.message || String(this.state.error)}
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}
