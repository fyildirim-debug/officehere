// Uzun dosya işlemleri için ilerleme göstergesi: "Açılıyor… %40 [İptal]". Durum çubuğunda ya da
// bir iletişim kutusunda kullanılır; oran verilmezse belirsiz (kayan) çubuk çizilir.
import { useId } from "react";
import { tipProps } from "./Tooltip";

export interface ProgressProps {
  /** Kısa durum metni ("Açılıyor…", "Kaydediliyor…"); çubuğun erişilebilir adıdır. */
  label: string;
  /** 0..1 arası oran; verilmezse (ya da geçersizse) belirsiz ilerleme. */
  value?: number | null;
  /** Adım adı ("Sayfalar okunuyor"): ipucunda ve ekran okuyucu değer metninde. */
  detail?: string;
  /** Verilirse İptal düğmesi gösterilir. */
  onCancel?: () => void;
  /** İptal düğmesinin metni (varsayılan "İptal"). */
  cancelLabel?: string;
  className?: string;
}

/** Oranı 0..100 tam sayı yüzdeye çevirir; geçersizse null (belirsiz). */
export function progressPercent(value: number | null | undefined): number | null {
  if (value == null || !Number.isFinite(value)) return null;
  return Math.round(Math.min(1, Math.max(0, value)) * 100);
}

export function Progress({ label, value, detail, onCancel, cancelLabel = "İptal", className }: ProgressProps) {
  const labelId = `${useId()}-etiket`;
  const pct = progressPercent(value);
  const text = pct === null ? undefined : `%${pct}`;
  const valueText = [text, detail].filter(Boolean).join(", ") || undefined;
  return (
    <div className={`oh-progress${pct === null ? " is-indeterminate" : ""}${className ? ` ${className}` : ""}`}>
      <span id={labelId} className="oh-progress-label">
        {label}
      </span>
      <span
        className="oh-progress-track"
        role="progressbar"
        aria-labelledby={labelId}
        aria-valuemin={0}
        aria-valuemax={100}
        aria-valuenow={pct ?? undefined}
        aria-valuetext={valueText}
        {...tipProps(detail)}
      >
        <span className="oh-progress-fill" style={pct === null ? undefined : { width: `${pct}%` }} />
      </span>
      {text && (
        <span className="oh-progress-pct" aria-hidden="true">
          {text}
        </span>
      )}
      {onCancel && (
        <button type="button" className="oh-progress-cancel" onClick={onCancel}>
          {cancelLabel}
        </button>
      )}
    </div>
  );
}
