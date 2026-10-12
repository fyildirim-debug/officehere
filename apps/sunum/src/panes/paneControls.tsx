/** Biçim bölmelerinde kullanılan küçük denetimler. */
import { useCallback, useRef, useState, type ReactNode } from "react";
import { MenuProvider, Popover } from "@officehere/ui";
import { cssColor } from "../model/color";
import type { Color, Theme } from "../model/types";
import { ThemePalette } from "../ribbon/common";

export function PaneColorButton({ theme, value, onChange, noneLabel }: { theme: Theme; value: Color | undefined; onChange: (c: Color | null) => void; noneLabel?: string }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLButtonElement>(null);
  const close = useCallback(() => setOpen(false), []);
  return (
    <>
      <button ref={ref} type="button" className="color-btn" onClick={() => setOpen(!open)} aria-label="Renk">
        <span className="color-btn-swatch" style={{ background: value ? cssColor(value, theme) : "transparent" }} />
        <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden="true">
          <path d="M1 2.5l3 3 3-3" fill="none" stroke="currentColor" strokeWidth="1.1" />
        </svg>
      </button>
      <Popover open={open} onClose={close} anchor={ref}>
        <MenuProvider onClose={close}>
          <div className="oh-menu oh-menu-palette">
            <ThemePalette theme={theme} value={value ?? null} noneLabel={noneLabel} onSelect={onChange} />
          </div>
        </MenuProvider>
      </Popover>
    </>
  );
}

export function Section({ title, children, open = true }: { title: string; children: ReactNode; open?: boolean }) {
  return (
    <details className="fp-section" open={open}>
      <summary>{title}</summary>
      {children}
    </details>
  );
}

export function Row({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="fp-row">
      <span>{label}</span>
      <span>{children}</span>
    </div>
  );
}

export function NumInput({ value, onChange, step = 1, min, max, suffix }: { value: number; onChange: (v: number) => void; step?: number; min?: number; max?: number; suffix?: string }) {
  const [text, setText] = useState<string | null>(null);
  const shown = text ?? `${Number(value.toFixed(2)).toLocaleString("tr-TR")}${suffix ? ` ${suffix}` : ""}`;
  const commit = (s: string) => {
    const n = parseFloat(s.replace(",", ".").replace(/[^\d.-]/g, ""));
    setText(null);
    if (Number.isFinite(n)) onChange(Math.min(max ?? Infinity, Math.max(min ?? -Infinity, n)));
  };
  return (
    <input
      type="text"
      value={shown}
      onFocus={(e) => setText(e.currentTarget.value)}
      onChange={(e) => setText(e.target.value)}
      onBlur={(e) => commit(e.target.value)}
      onKeyDown={(e) => {
        e.stopPropagation();
        if (e.key === "Enter") commit((e.target as HTMLInputElement).value);
        if (e.key === "ArrowUp") (e.preventDefault(), onChange(Math.min(max ?? Infinity, value + step)));
        if (e.key === "ArrowDown") (e.preventDefault(), onChange(Math.max(min ?? -Infinity, value - step)));
      }}
    />
  );
}

export function Slider({ value, onChange, min = 0, max = 100 }: { value: number; onChange: (v: number) => void; min?: number; max?: number }) {
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
      <input type="range" min={min} max={max} value={value} onChange={(e) => onChange(Number(e.target.value))} />
      <span style={{ width: 34, textAlign: "right" }}>%{Math.round(value)}</span>
    </span>
  );
}

export function Radio({ label, checked, onChange }: { label: string; checked: boolean; onChange: () => void }) {
  return (
    <label className="fp-radio">
      <input type="radio" checked={checked} onChange={onChange} />
      {label}
    </label>
  );
}

export function PaneShell({ title, onClose, children, tabs }: { title: string; onClose: () => void; children: ReactNode; tabs?: ReactNode }) {
  return (
    <aside className="side-pane" onKeyDown={(e) => e.stopPropagation()}>
      <div className="side-pane-head">
        <span>{title}</span>
        <button type="button" className="side-pane-close" onClick={onClose} aria-label="Kapat" title="Kapat">
          ✕
        </button>
      </div>
      {tabs}
      <div className="side-pane-body oh-scroll">{children}</div>
    </aside>
  );
}
