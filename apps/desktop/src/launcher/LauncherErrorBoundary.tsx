// Başlangıç ekranının hata sınırı: bir bileşen çökerse boş/donuk pencere yerine
// anlaşılır bir ileti, "Yeniden dene" ve "Hata bildir" gösterilir. Hata yerel günlüğe yazılır.
import { Component, type ErrorInfo, type ReactNode } from "react";
import { BugReportDialog, logPageEvent } from "@officehere/ui";

interface State {
  error: Error | null;
  key: number;
  reporting: boolean;
}

export class LauncherErrorBoundary extends Component<{ children: ReactNode }, State> {
  state: State = { error: null, key: 0, reporting: false };

  static getDerivedStateFromError(error: Error): Partial<State> {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    logPageEvent({ level: "hata", message: `Başlangıç ekranı: ${error.name}: ${error.message}`, stack: error.stack ?? info.componentStack ?? undefined, kind: "react" });
  }

  render() {
    const { error, key, reporting } = this.state;
    if (!error) return <div key={key} style={{ display: "contents" }}>{this.props.children}</div>;
    return (
      <div className="ohl-crash" role="alert">
        <div className="ohl-crash-box">
          <strong>Başlangıç ekranında bir sorun oluştu</strong>
          <p>Açık belgeleriniz etkilenmedi. Ekranı yeniden yükleyebilir ya da sorunu bildirebilirsiniz.</p>
          <p className="ohl-crash-detail">{error.message}</p>
          <div className="ohl-crash-actions">
            <button type="button" className="oh-btn is-primary" onClick={() => this.setState((s) => ({ error: null, key: s.key + 1 }))}>
              Yeniden dene
            </button>
            <button type="button" className="oh-btn" onClick={() => this.setState({ reporting: true })}>
              Hata bildir
            </button>
          </div>
        </div>
        <BugReportDialog open={reporting} onClose={() => this.setState({ reporting: false })} app="baslangic" />
      </div>
    );
  }
}
