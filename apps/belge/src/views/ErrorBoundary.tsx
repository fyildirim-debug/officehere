// Hata sınırları: bir bileşen çökerse belge kaybolmaz. Hata anında otomatik kurtarma taslağı yazılır;
// bölme yerinde "Bir sorun oluştu" uyarısı ve Belgeyi Kurtar / Kaydet seçenekleri gösterilir.
import { Component, useEffect, useState, type ErrorInfo, type ReactNode } from "react";
import { baseName, downloadBlob } from "@officehere/core";
import { writeDocx } from "../io";
import { writeDraft, type Draft } from "../state/drafts";

/** Kök hata sınırından "Belgeyi Kurtar": yeniden kurulan uygulama kendi taslağını yükler (aynı sayfa). */
let autoRestore = false;

/** Uygulamanın kayıt işlevleri (App her çizimde günceller). */
export const recovery: { snapshot: (() => Draft | null) | null; save: (() => Promise<unknown>) | null } = { snapshot: null, save: null };

/** Güncel belgeyi otomatik kurtarma taslağı olarak yazar. */
export function saveRecoveryDraft(): Draft | null {
  let d: Draft | null = null;
  try {
    d = recovery.snapshot?.() ?? null;
    // Taslak eşzamansız (IndexedDB) yazılır; sonraki okumalar yazmanın ardından sıraya girer.
    if (d) void writeDraft(d);
  } catch {
    // depolama kapalı ya da belge okunamadı
  }
  return d;
}

/** Açılışta taslak kendiliğinden geri yüklensin mi (kök hata sınırından "Belgeyi Kurtar"). */
export function takeAutoRestore(): boolean {
  const v = autoRestore;
  autoRestore = false;
  return v;
}

async function downloadDraft(d: Draft | null) {
  if (!d) return;
  const bytes = await writeDocx({ content: d.content, settings: d.settings });
  downloadBlob(new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }), `${baseName(d.fileName)} (kurtarılan).docx`);
}

/** Test ve tanılama için: `officehere:crash-test` olayı adı verilen bölmeyi bir kez çökertir. */
function CrashProbe({ name }: { name: string }) {
  const [crash, setCrash] = useState(false);
  useEffect(() => {
    const on = (e: Event) => {
      if ((e as CustomEvent<string>).detail === name) setCrash(true);
    };
    window.addEventListener("officehere:crash-test", on);
    return () => window.removeEventListener("officehere:crash-test", on);
  }, [name]);
  if (crash) throw new Error(`Deneme hatası (${name})`);
  return null;
}

interface Props {
  /** Bölme adı (uyarıda gösterilir; çökme denemesi için anahtar). */
  name: string;
  label: string;
  children: ReactNode;
  /** Kök sınır: tüm uygulama yerine tam sayfa kurtarma ekranı. */
  root?: boolean;
  /** Kurtarmada (yeniden kurmadan önce) çağrılır; ör. çöken iletişim kutusunu kapatmak için. */
  onReset?: () => void;
}

interface State {
  error: Error | null;
  key: number;
}

export class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null, key: 0 };

  static getDerivedStateFromError(error: Error): Partial<State> {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    saveRecoveryDraft();
    console.warn(`[OfficeHere Belge] ${this.props.label} hatası:`, error, info.componentStack);
  }

  private reset = () => {
    this.props.onReset?.();
    this.setState((s) => ({ error: null, key: s.key + 1 }));
  };

  private recover = () => {
    saveRecoveryDraft();
    if (this.props.root) autoRestore = true;
    this.reset();
  };

  private save = async () => {
    if (!this.props.root && recovery.save) {
      await recovery.save();
      return;
    }
    await downloadDraft(saveRecoveryDraft());
  };

  render() {
    const { error, key } = this.state;
    if (!error) {
      return (
        <ErrorScope key={key}>
          <CrashProbe name={this.props.name} />
          {this.props.children}
        </ErrorScope>
      );
    }
    return (
      <div className={`bl-crash${this.props.root ? " is-root" : ""}`} role="alert">
        <div className="bl-crash-box">
          <strong>Bir sorun oluştu — belgeyi kurtar/kaydet</strong>
          <p>
            {this.props.label} beklenmedik bir hatayla karşılaştı. Belgeniz otomatik kurtarma taslağı olarak saklandı; çalışmaya devam etmek için belgeyi kurtarın ya da bir
            kopyasını kaydedin.
          </p>
          <p className="bl-crash-detail">{error.message}</p>
          <div className="bl-crash-actions">
            <button type="button" className="oh-btn is-primary" onClick={this.recover}>
              Belgeyi Kurtar
            </button>
            <button type="button" className="oh-btn" onClick={() => void this.save()}>
              Kaydet
            </button>
          </div>
        </div>
      </div>
    );
  }
}

/** Sıfırlamada alt ağacın yeniden kurulması için anahtarlı sarmalayıcı. */
function ErrorScope({ children }: { children: ReactNode }) {
  return <>{children}</>;
}
