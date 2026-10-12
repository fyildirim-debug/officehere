/**
 * Hata sınırı: bir bölge (şerit, formül çubuğu, çalışma sayfası, bölmeler,
 * dosya görünümü) çizim sırasında çökerse yalnız o bölge bir uyarıyla
 * değiştirilir. Çalışma kitabı denetleyicide (bellekte) kaldığı için
 * kaybolmaz: uyarıdan kaydedilebilir, kurtarma kopyası indirilebilir ya da
 * bölge yeniden denenebilir.
 */
import { Component, type ErrorInfo, type ReactNode } from "react";

export interface RecoveryHandlers {
  save: () => void;
  download: () => void;
  undo?: () => void;
}

interface Props {
  region: string;
  recovery: RecoveryHandlers;
  children: ReactNode;
  /** Kök sınırı (tüm uygulama). */
  root?: boolean;
  /** Küçük bölgeler için tek satırlık uyarı. */
  compact?: boolean;
}

interface State {
  error: Error | null;
  attempt: number;
}

/** Tanılama ve testler için: window.__ohTabloCrash = "<bölge>" o bölgeyi bir kez çökertir. */
function CrashProbe({ region }: { region: string }) {
  const w = window as unknown as { __ohTabloCrash?: string };
  if (w.__ohTabloCrash === region) {
    window.setTimeout(() => {
      if (w.__ohTabloCrash === region) w.__ohTabloCrash = undefined;
    }, 0);
    throw new Error(`Tanılama amaçlı çökme: ${region}`);
  }
  return null;
}

export class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null, attempt: 0 };

  static getDerivedStateFromError(error: Error): Partial<State> {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    console.warn(`[OfficeHere Tablo] ${this.props.region} çöktü:`, error.message, info.componentStack?.split("\n").slice(0, 4).join("\n"));
  }

  private retry = (undo: boolean) => {
    if (undo) this.props.recovery.undo?.();
    this.setState((s) => ({ error: null, attempt: s.attempt + 1 }));
  };

  render() {
    const { error, attempt } = this.state;
    if (!error) {
      return (
        <Scope key={attempt}>
          <CrashProbe region={this.props.region} />
          {this.props.children}
        </Scope>
      );
    }
    const r = this.props.recovery;
    return (
      <div className={`tb-crash${this.props.root ? " is-root" : ""}${this.props.compact ? " is-compact" : ""}`} role="alert" data-region={this.props.region}>
        <div className="tb-crash-title">Bir sorun oluştu — {this.props.region} görüntülenemedi.</div>
        {!this.props.compact && <div className="tb-crash-text">Çalışma kitabınız kaybolmadı. Hemen kaydedebilir, bir kurtarma kopyası indirebilir ya da yeniden deneyebilirsiniz.</div>}
        <div className="tb-crash-actions">
          <button type="button" className="oh-btn is-primary" onClick={r.save}>
            Kaydet
          </button>
          <button type="button" className="oh-btn" onClick={r.download}>
            Kurtarma Kopyası İndir
          </button>
          <button type="button" className="oh-btn" onClick={() => this.retry(false)}>
            Yeniden Dene
          </button>
          {r.undo && (
            <button type="button" className="oh-btn" onClick={() => this.retry(true)}>
              Geri Al ve Yeniden Dene
            </button>
          )}
        </div>
        {!this.props.compact && (
          <details className="tb-crash-details">
            <summary>Ayrıntılar</summary>
            <code>{error.message}</code>
          </details>
        )}
      </div>
    );
  }
}

function Scope({ children }: { children: ReactNode }) {
  return <>{children}</>;
}
