/** İletişim kutuları için küçük ortak denetimler. */
import { Component, useRef, useState, type ReactNode } from "react";
import { ColorPalette, Dialog, DialogButton, MenuProvider, Popover } from "@officehere/ui";

export interface BaseDialogProps {
  onClose: () => void;
}

export function OkCancel({ onOk, onCancel, okLabel = "Tamam", disabled }: { onOk: () => void; onCancel: () => void; okLabel?: string; disabled?: boolean }) {
  return (
    <>
      <DialogButton primary onClick={onOk} disabled={disabled}>
        {okLabel}
      </DialogButton>
      <DialogButton onClick={onCancel}>İptal</DialogButton>
    </>
  );
}

export function SimpleDialog({ title, onClose, onOk, children, width = 360, okLabel, disabled }: { title: string; onClose: () => void; onOk: () => void; children: ReactNode; width?: number; okLabel?: string; disabled?: boolean }) {
  return (
    <Dialog open title={title} onClose={onClose} width={width} onSubmit={disabled ? undefined : onOk} footer={<OkCancel onOk={onOk} onCancel={onClose} okLabel={okLabel} disabled={disabled} />}>
      {children}
    </Dialog>
  );
}

export function Radio<T extends string>({ name, value, current, onChange, children }: { name: string; value: T; current: T; onChange: (v: T) => void; children: ReactNode }) {
  return (
    <label className="tb-radio">
      <input type="radio" name={name} checked={current === value} onChange={() => onChange(value)} />
      <span>{children}</span>
    </label>
  );
}

export function Check({ checked, onChange, children }: { checked: boolean; onChange: (v: boolean) => void; children: ReactNode }) {
  return (
    <label className="tb-check">
      <input type="checkbox" checked={checked} onChange={(e) => onChange(e.target.checked)} />
      <span>{children}</span>
    </label>
  );
}

/** Türkçe ondalık virgüllü sayı ayrıştırma. */
export function parseTrNumber(s: string): number {
  return Number(s.trim().replace(/\s/g, "").replace(/\.(?=\d{3}(\D|$))/g, "").replace(",", "."));
}

export function fmtTr(n: number, digits = 2): string {
  return Number(n.toFixed(digits)).toLocaleString("tr-TR", { maximumFractionDigits: digits });
}

/** İletişim kutusu içinde renk seçici açılır düğmesi. */
export function ColorDrop({ value, onChange, noneLabel = "Otomatik", width = 120 }: { value: string | null | undefined; onChange: (c: string | null) => void; noneLabel?: string; width?: number }) {
  const ref = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  return (
    <>
      <button ref={ref} type="button" className="tb-colordrop" style={{ width }} onClick={() => setOpen((o) => !o)} aria-label="Renk">
        <span className="tb-colordrop-swatch" style={{ background: value ?? "transparent" }} />
        <span className="tb-colordrop-label">{value ? value.toUpperCase() : noneLabel}</span>
        <span className="tb-colordrop-arrow">▾</span>
      </button>
      <Popover open={open} onClose={() => setOpen(false)} anchor={ref} className="tb-dialog-pop" ignore={[ref]}>
        <MenuProvider onClose={() => setOpen(false)}>
          <div className="oh-menu">
            <ColorPalette value={value ?? null} noneLabel={noneLabel} onSelect={(c) => (onChange(c), setOpen(false))} />
          </div>
        </MenuProvider>
      </Popover>
    </>
  );
}

/** İletişim kutusu çizimindeki beklenmeyen hata tüm uygulamayı düşürmesin. */
export class DialogErrorBoundary extends Component<{ children: ReactNode; onError: (e: Error) => void; resetKey: unknown }, { failed: boolean; key: unknown }> {
  state = { failed: false, key: undefined as unknown };

  static getDerivedStateFromProps(props: { resetKey: unknown }, state: { failed: boolean; key: unknown }) {
    return props.resetKey !== state.key ? { failed: false, key: props.resetKey } : null;
  }

  static getDerivedStateFromError() {
    return { failed: true };
  }

  componentDidCatch(error: Error) {
    this.props.onError(error);
  }

  render() {
    return this.state.failed ? null : this.props.children;
  }
}
