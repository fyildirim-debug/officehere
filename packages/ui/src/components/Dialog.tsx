import { useEffect, useRef, useState, type ReactNode } from "react";
import { createPortal } from "react-dom";

export interface DialogProps {
  open: boolean;
  title: string;
  onClose: () => void;
  children: ReactNode;
  footer?: ReactNode;
  width?: number;
  /** Enter tuşunun tetikleyeceği eylem. */
  onSubmit?: () => void;
}

/** Sürüklenebilir, modal ofis iletişim kutusu. */
export function Dialog({ open, title, onClose, children, footer, width = 420, onSubmit }: DialogProps) {
  const [offset, setOffset] = useState({ x: 0, y: 0 });
  const boxRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (open) setOffset({ x: 0, y: 0 });
  }, [open]);

  useEffect(() => {
    if (!open) return;
    const prev = document.activeElement as HTMLElement | null;
    const first = boxRef.current?.querySelector<HTMLElement>("input, select, textarea, button:not(.oh-dialog-close)");
    first?.focus();
    return () => prev?.focus?.();
  }, [open]);

  if (!open) return null;

  const startDrag = (e: React.PointerEvent) => {
    if ((e.target as HTMLElement).closest("button")) return;
    const sx = e.clientX - offset.x;
    const sy = e.clientY - offset.y;
    const move = (ev: PointerEvent) => setOffset({ x: ev.clientX - sx, y: ev.clientY - sy });
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };

  return createPortal(
    <div className="oh-dialog-backdrop" onMouseDown={(e) => e.target === e.currentTarget && e.preventDefault()}>
      <div
        ref={boxRef}
        className="oh-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={title}
        style={{ width, transform: `translate(${offset.x}px, ${offset.y}px)` }}
        onKeyDown={(e) => {
          if (e.key === "Escape") {
            e.stopPropagation();
            onClose();
          } else if (e.key === "Enter" && onSubmit && (e.target as HTMLElement).tagName !== "TEXTAREA" && (e.target as HTMLElement).tagName !== "BUTTON") {
            e.preventDefault();
            onSubmit();
          }
        }}
      >
        <div className="oh-dialog-title" onPointerDown={startDrag}>
          <span>{title}</span>
          <button type="button" className="oh-dialog-close" onClick={onClose} aria-label="Kapat">
            <svg width="10" height="10" viewBox="0 0 10 10" aria-hidden="true">
              <path d="M1 1l8 8M9 1L1 9" stroke="currentColor" strokeWidth="1.2" />
            </svg>
          </button>
        </div>
        <div className="oh-dialog-body">{children}</div>
        {footer && <div className="oh-dialog-footer">{footer}</div>}
      </div>
    </div>,
    document.body,
  );
}

export function DialogButton({ children, primary, onClick, disabled }: { children: ReactNode; primary?: boolean; onClick: () => void; disabled?: boolean }) {
  return (
    <button type="button" className={`oh-btn${primary ? " is-primary" : ""}`} onClick={onClick} disabled={disabled}>
      {children}
    </button>
  );
}

/** İletişim kutusu içinde etiket + denetim satırı. */
export function Field({ label, children, inline }: { label: string; children: ReactNode; inline?: boolean }) {
  return (
    <label className={`oh-field${inline ? " is-inline" : ""}`}>
      <span className="oh-field-label">{label}</span>
      {children}
    </label>
  );
}

export function FieldSet({ legend, children }: { legend: string; children: ReactNode }) {
  return (
    <fieldset className="oh-fieldset">
      <legend>{legend}</legend>
      {children}
    </fieldset>
  );
}

/** İletişim kutusu sekmeleri (Paragraf: "Girintiler ve Aralıklar" / "Satır ve Sayfa Sonları"). */
export function DialogTabs({ tabs, active, onChange }: { tabs: { id: string; label: string }[]; active: string; onChange: (id: string) => void }) {
  return (
    <div className="oh-dialog-tabs" role="tablist">
      {tabs.map((t) => (
        <button key={t.id} type="button" role="tab" aria-selected={t.id === active} className={t.id === active ? "is-active" : ""} onClick={() => onChange(t.id)}>
          {t.label}
        </button>
      ))}
    </div>
  );
}

export type ConfirmChoice = "save" | "discard" | "cancel";

/** "Değişiklikleri kaydetmek istiyor musunuz?" iletişim kutusu. */
export function SaveChangesDialog({ open, fileName, onChoice }: { open: boolean; fileName: string; onChoice: (c: ConfirmChoice) => void }) {
  return (
    <Dialog
      open={open}
      title="OfficeHere"
      onClose={() => onChoice("cancel")}
      width={440}
      footer={
        <>
          <DialogButton primary onClick={() => onChoice("save")}>
            Kaydet
          </DialogButton>
          <DialogButton onClick={() => onChoice("discard")}>Kaydetme</DialogButton>
          <DialogButton onClick={() => onChoice("cancel")}>İptal</DialogButton>
        </>
      }
    >
      <p className="oh-dialog-message">
        <strong>"{fileName}"</strong> dosyasındaki değişiklikleri kaydetmek istiyor musunuz?
      </p>
    </Dialog>
  );
}

/** Şeridin altında görünen bilgi çubuğu. */
export function MessageBar({
  children,
  kind = "info",
  onClose,
  actions,
  closeLabel = "Kapat",
}: {
  children: ReactNode;
  kind?: "info" | "warning" | "error" | "success";
  onClose?: () => void;
  actions?: ReactNode;
  /** Kapat düğmesinin adı ve ipucu (kapatmanın bir sonucu varsa söylenir, ör. "Taslağı sil ve kapat"). */
  closeLabel?: string;
}) {
  return (
    <div className={`oh-messagebar is-${kind}`} role="status">
      <span className="oh-messagebar-text">{children}</span>
      {actions}
      {onClose && (
        <button type="button" className="oh-messagebar-close" onClick={onClose} aria-label={closeLabel} title={closeLabel}>
          <svg width="10" height="10" viewBox="0 0 10 10" aria-hidden="true">
            <path d="M1 1l8 8M9 1L1 9" stroke="currentColor" strokeWidth="1.2" />
          </svg>
        </button>
      )}
    </div>
  );
}
