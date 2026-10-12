/**
 * Hata sınırı: bir bölme (slayt alanı, küçük resimler, bölmeler, gösteri)
 * çizim sırasında çökerse yalnız o bölge bir uyarıyla değiştirilir; sunu
 * bellekteki depoda (Store) kaldığı için kaybolmaz. Uyarıdan belge
 * kaydedilebilir, kurtarma kopyası indirilebilir ya da bölge yeniden
 * denenebilir (gerekirse son değişiklik geri alınarak).
 */
import { Component, type ErrorInfo, type ReactNode } from "react";

export interface RecoveryHandlers {
  /** Belgeyi normal yoldan kaydet (Ctrl+S). */
  save: () => void;
  /** Sunuyu .pptx olarak indir (kayıt yolu çalışmasa bile). */
  download: () => void;
  /** Son değişikliği geri al (çökmeye yol açan adım). */
  undo?: () => void;
}

interface Props {
  /** Bölge adı ("Slayt alanı" vb.). */
  region: string;
  recovery: RecoveryHandlers;
  children: ReactNode;
  /** Kök sınırı: tam sayfa uyarı. */
  root?: boolean;
  /** Bölge yeniden denendiğinde. */
  onReset?: () => void;
}

interface State {
  error: Error | null;
  attempt: number;
}

/** Testler ve tanılama için: window.__ohSunumCrash = "<bölge>" o bölgeyi bir kez çökertir. */
function CrashProbe({ region }: { region: string }) {
  const w = window as unknown as { __ohSunumCrash?: string };
  if (w.__ohSunumCrash === region) {
    w.__ohSunumCrash = undefined;
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
    // Ayrıntı geliştirici konsoluna; kullanıcıya sade uyarı gösterilir.
    console.warn(`[OfficeHere Sunum] ${this.props.region} çöktü:`, error.message, info.componentStack?.split("\n").slice(0, 4).join("\n"));
  }

  private reset = (undo: boolean) => {
    if (undo) this.props.recovery.undo?.();
    this.props.onReset?.();
    this.setState((s) => ({ error: null, attempt: s.attempt + 1 }));
  };

  render() {
    const { error, attempt } = this.state;
    if (!error) {
      return (
        <ErrorScope key={attempt}>
          <CrashProbe region={this.props.region} />
          {this.props.children}
        </ErrorScope>
      );
    }
    const r = this.props.recovery;
    return (
      <div className={`crash-box${this.props.root ? " is-root" : ""}`} role="alert" data-region={this.props.region}>
        <div className="crash-title">Bir sorun oluştu — {this.props.region} görüntülenemedi.</div>
        <div className="crash-text">Sununuz kaybolmadı. Çalışmanızı hemen kaydedebilir ya da bir kurtarma kopyası indirebilirsiniz.</div>
        <div className="crash-actions">
          <button type="button" className="oh-btn is-primary" onClick={r.save}>
            Belgeyi Kaydet
          </button>
          <button type="button" className="oh-btn" onClick={r.download}>
            Kurtarma Kopyası İndir
          </button>
          <button type="button" className="oh-btn" onClick={() => this.reset(false)}>
            Yeniden Dene
          </button>
          {r.undo && (
            <button type="button" className="oh-btn" onClick={() => this.reset(true)}>
              Geri Al ve Yeniden Dene
            </button>
          )}
        </div>
        <details className="crash-details">
          <summary>Ayrıntılar</summary>
          <code>{error.message}</code>
        </details>
      </div>
    );
  }
}

/** Yeniden denemede alt ağacı sıfırdan kurmak için anahtar taşıyıcı. */
function ErrorScope({ children }: { children: ReactNode }) {
  return <>{children}</>;
}
