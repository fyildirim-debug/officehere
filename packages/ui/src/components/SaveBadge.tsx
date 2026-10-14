// Başlık çubuğundaki kayıt durumu rozeti: Kaydedilmedi / Kaydediliyor… / Kaydedildi HH:MM /
// Kaydedilemedi: <neden> / Otomatik kayıt duraklatıldı.
import { Checkmark12Filled, ErrorCircle12Regular, Pause12Regular } from "@fluentui/react-icons";
import { tipProps } from "./Tooltip";
import { Dialog, DialogButton } from "./Dialog";
import { formatSavedTime, saveStatus, useSaveAlert, type SaveState } from "../hooks/saveStatus";

export interface SaveBadgeProps {
  state: SaveState;
  savedAt?: number;
  /** "Kaydedilemedi" tıklanınca (verilmezse kaydetme akışının kaydettiği Kaydet). */
  onRetry?: () => void;
  /** Kaydetme hatasının kısa nedeni ("Disk dolu"). */
  error?: string;
  /** Hatanın uzun iletisi (ipucunda). */
  errorDetail?: string;
  /** Otomatik kaydetmenin duraklatılma nedeni. */
  pausedReason?: string;
  /** "hold": ilk elle kaydı bekliyor; "off": bu dosyada kapalı. */
  pausedKind?: "hold" | "off";
}

export function SaveBadge({ state, savedAt, onRetry, error, errorDetail, pausedReason, pausedKind = "hold" }: SaveBadgeProps) {
  switch (state) {
    case "dirty":
      // oh-titlebar-dirty: masaüstü testleri ve eski seçiciler için korunur.
      return (
        <span className="oh-savebadge is-dirty oh-titlebar-dirty" {...tipProps("Kaydedilmemiş değişiklikler var")}>
          <span className="oh-savebadge-dot" aria-hidden="true" />
          Kaydedilmedi
        </span>
      );
    case "paused": {
      // Sakin bir dille: kullanıcıdan yalnız ilk kaydı onaylaması istenir (tıklanınca Kaydet).
      const hold = pausedKind === "hold";
      const label = hold ? "Otomatik kayıt bekliyor" : "Otomatik kayıt kapalı";
      const hint = hold
        ? "Kaydet (Ctrl+S) ile ilk kaydı onayladığınızda otomatik kaydetme sürer."
        : "Bu dosya için otomatik kaydetme kapatıldı; Kaydet ya da Farklı Kaydet ile kendiniz kaydedin.";
      return (
        <button
          type="button"
          className={`oh-savebadge is-paused oh-titlebar-dirty${hold ? "" : " is-off"}`}
          onClick={onRetry ?? saveStatus.retry}
          aria-label={`Kaydedilmedi. ${label}: ${pausedReason ?? ""}`}
          {...tipProps(hold ? "Otomatik kaydetme ilk kaydınızı bekliyor" : "Otomatik kaydetme kapalı", undefined, `${pausedReason ? `${pausedReason}. ` : ""}${hint}`)}
        >
          <Pause12Regular aria-hidden="true" />
          {label}
        </button>
      );
    }
    case "saving":
      return (
        <span className="oh-savebadge is-saving" {...tipProps("Belge kaydediliyor")}>
          <svg className="oh-savebadge-spin" width="12" height="12" viewBox="0 0 12 12" aria-hidden="true">
            <circle cx="6" cy="6" r="4.5" fill="none" stroke="currentColor" strokeOpacity="0.3" strokeWidth="1.4" />
            <path d="M6 1.5a4.5 4.5 0 0 1 4.5 4.5" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
          </svg>
          Kaydediliyor…
        </span>
      );
    case "saved": {
      const time = savedAt ? formatSavedTime(savedAt) : "";
      const full = savedAt ? new Date(savedAt).toLocaleString("tr-TR", { dateStyle: "short", timeStyle: "medium" }) : "";
      // key: her kayıtta sönükleşme hareketi baştan başlar.
      return (
        <span key={savedAt} className="oh-savebadge is-saved" {...tipProps("Tüm değişiklikler kaydedildi", undefined, full ? `Son kaydetme: ${full}` : undefined)}>
          <Checkmark12Filled aria-hidden="true" />
          {time ? `Kaydedildi ${time}` : "Kaydedildi"}
        </span>
      );
    }
    case "error":
      return (
        <button
          type="button"
          className="oh-savebadge is-error"
          onClick={onRetry ?? saveStatus.retry}
          {...tipProps(error ? `Kaydedilemedi: ${error}` : "Kaydedilemedi", undefined, `${errorDetail ? `${errorDetail} ` : ""}Yeniden kaydetmek için tıklayın.`)}
        >
          <ErrorCircle12Regular aria-hidden="true" />
          {error ? `Kaydedilemedi: ${error}` : "Kaydedilemedi"}
        </button>
      );
  }
}

/** Kayıtla ilgili önemli uyarı (ör. otomatik kaydedilen dosya doğrulanamadı): bir kez, onaylanınca kapanır. */
export function SaveAlertDialog() {
  const alert = useSaveAlert();
  const close = () => saveStatus.alert(null);
  return (
    <Dialog open={!!alert} title={alert?.title ?? ""} onClose={close} width={460} footer={<DialogButton primary onClick={close}>Tamam</DialogButton>}>
      <p className="oh-dialog-message" role="alert">
        {alert?.message}
      </p>
    </Dialog>
  );
}
