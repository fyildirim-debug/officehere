// Başlık çubuğundaki kayıt durumu rozeti: Kaydedilmedi / Kaydediliyor… / Kaydedildi HH:MM / Kaydedilemedi.
import { Checkmark12Filled, ErrorCircle12Regular } from "@fluentui/react-icons";
import { tipProps } from "./Tooltip";
import { formatSavedTime, saveStatus, type SaveState } from "../hooks/saveStatus";

export interface SaveBadgeProps {
  state: SaveState;
  savedAt?: number;
  /** "Kaydedilemedi" tıklanınca (verilmezse kaydetme akışının kaydettiği Kaydet). */
  onRetry?: () => void;
}

export function SaveBadge({ state, savedAt, onRetry }: SaveBadgeProps) {
  switch (state) {
    case "dirty":
      // oh-titlebar-dirty: masaüstü testleri ve eski seçiciler için korunur.
      return (
        <span className="oh-savebadge is-dirty oh-titlebar-dirty" {...tipProps("Kaydedilmemiş değişiklikler var")}>
          <span className="oh-savebadge-dot" aria-hidden="true" />
          Kaydedilmedi
        </span>
      );
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
        <button type="button" className="oh-savebadge is-error" onClick={onRetry ?? saveStatus.retry} {...tipProps("Kaydedilemedi", undefined, "Yeniden kaydetmek için tıklayın.")}>
          <ErrorCircle12Regular aria-hidden="true" />
          Kaydedilemedi
        </button>
      );
  }
}
