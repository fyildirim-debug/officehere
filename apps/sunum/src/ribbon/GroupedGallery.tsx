/**
 * Kategorili şerit galerisi (Geçişler: Hafif / Heyecan Verici …). Ortak
 * RibbonGallery ile aynı görünüm; açılınca öğeler kategori başlıklarıyla listelenir.
 */
import { useCallback, useRef, useState, type ReactNode } from "react";
import { MenuProvider, Popover, tipProps } from "@officehere/ui";

export interface GroupedItem {
  id: string;
  label: string;
  group: string;
  preview: ReactNode;
}

export function GroupedGallery({ items, selectedId, onSelect, itemWidth = 64, itemHeight = 58, columns = 8, expandedColumns = 6 }: { items: GroupedItem[]; selectedId?: string; onSelect: (id: string) => void; itemWidth?: number; itemHeight?: number; columns?: number; expandedColumns?: number }) {
  const [row, setRow] = useState(0);
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const close = useCallback(() => setOpen(false), []);
  const rows = Math.max(1, Math.ceil(items.length / columns));
  const visible = items.slice(row * columns, row * columns + columns);
  const groups = Array.from(new Set(items.map((i) => i.group)));

  const tile = (it: GroupedItem, inMenu: boolean) => (
    <button
      key={it.id}
      type="button"
      className={`oh-gallery-item${it.id === selectedId ? " is-selected" : ""}`}
      style={{ width: itemWidth, height: itemHeight }}
      aria-label={it.label}
      aria-pressed={it.id === selectedId}
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => {
        onSelect(it.id);
        if (inMenu) close();
      }}
      {...tipProps(it.label)}
    >
      <span className="oh-gallery-preview">{it.preview}</span>
    </button>
  );

  return (
    <div ref={ref} className="oh-gallery">
      <div className="oh-gallery-items" style={{ width: columns * (itemWidth + 2) + 2, height: itemHeight + 4 }}>
        {visible.map((it) => tile(it, false))}
      </div>
      <div className="oh-gallery-scroll">
        <button type="button" disabled={row === 0} onMouseDown={(e) => e.preventDefault()} onClick={() => setRow((r) => Math.max(0, r - 1))} aria-label="Yukarı">
          <svg width="8" height="5" viewBox="0 0 8 5" aria-hidden="true">
            <path d="M0 5l4-5 4 5z" fill="currentColor" />
          </svg>
        </button>
        <button type="button" disabled={row >= rows - 1} onMouseDown={(e) => e.preventDefault()} onClick={() => setRow((r) => Math.min(rows - 1, r + 1))} aria-label="Aşağı">
          <svg width="8" height="5" viewBox="0 0 8 5" aria-hidden="true">
            <path d="M0 0l4 5 4-5z" fill="currentColor" />
          </svg>
        </button>
        <button type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen(true)} aria-label="Diğer" {...tipProps("Diğer")}>
          <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden="true">
            <path d="M0 1h8" stroke="currentColor" />
            <path d="M0 3l4 5 4-5z" fill="currentColor" />
          </svg>
        </button>
      </div>
      <Popover open={open} onClose={close} anchor={ref}>
        <MenuProvider onClose={close}>
          <div className="oh-menu oh-gallery-menu">
            {groups.map((g) => (
              <div key={g} className="gallery-group" role="group" aria-label={g}>
                <div className="oh-menu-header">{g}</div>
                <div className="oh-gallery-grid" style={{ gridTemplateColumns: `repeat(${expandedColumns}, ${itemWidth}px)` }}>
                  {items.filter((i) => i.group === g).map((it) => tile(it, true))}
                </div>
              </div>
            ))}
          </div>
        </MenuProvider>
      </Popover>
    </div>
  );
}

/** "01,00" (saniye) ya da "00:05,00" (dakika:saniye) biçimli süre kutusu. */
export function TimeSpinner({ label, value, onChange, minutes, step = 0.25, max = 3600, tip, width = 72, disabled }: { label?: string; value: number; onChange: (v: number) => void; minutes?: boolean; step?: number; max?: number; tip: string; width?: number; disabled?: boolean }) {
  const fmt = (v: number) => {
    const cs = Math.round(v * 100);
    const sec = Math.floor(cs / 100);
    const frac = String(cs % 100).padStart(2, "0");
    if (!minutes) return `${String(sec).padStart(2, "0")},${frac}`;
    return `${String(Math.floor(sec / 60)).padStart(2, "0")}:${String(sec % 60).padStart(2, "0")},${frac}`;
  };
  const [text, setText] = useState(fmt(value));
  const [last, setLast] = useState(value);
  if (last !== value) {
    setLast(value);
    setText(fmt(value));
  }
  const clamp = (v: number) => Math.min(max, Math.max(0, Math.round(v * 100) / 100));
  const parse = (t: string): number | null => {
    const m = t.trim().replace(",", ".").match(/^(?:(\d+):)?(\d+(?:\.\d+)?)$/);
    if (!m) return null;
    return (m[1] ? Number(m[1]) * 60 : 0) + Number(m[2]);
  };
  const commit = () => {
    const n = parse(text);
    if (n !== null) onChange(clamp(n));
    else setText(fmt(value));
  };
  return (
    <div className="oh-spinner-wrap" {...tipProps(tip)}>
      {label && <span className="oh-spinner-label">{label}</span>}
      <div className="oh-spinner" style={{ width }}>
        <input
          value={text}
          disabled={disabled}
          onChange={(e) => setText(e.target.value)}
          onBlur={commit}
          onKeyDown={(e) => {
            if (e.key === "Enter") commit();
            if (e.key === "ArrowUp") (e.preventDefault(), onChange(clamp(value + step)));
            if (e.key === "ArrowDown") (e.preventDefault(), onChange(clamp(value - step)));
          }}
          aria-label={tip}
        />
        <div className="oh-spinner-btns">
          <button type="button" tabIndex={-1} disabled={disabled} onMouseDown={(e) => e.preventDefault()} onClick={() => onChange(clamp(value + step))} aria-label="Artır">
            <svg width="7" height="4" viewBox="0 0 7 4" aria-hidden="true">
              <path d="M0 4l3.5-4L7 4z" fill="currentColor" />
            </svg>
          </button>
          <button type="button" tabIndex={-1} disabled={disabled} onMouseDown={(e) => e.preventDefault()} onClick={() => onChange(clamp(value - step))} aria-label="Azalt">
            <svg width="7" height="4" viewBox="0 0 7 4" aria-hidden="true">
              <path d="M0 0l3.5 4L7 0z" fill="currentColor" />
            </svg>
          </button>
        </div>
      </div>
    </div>
  );
}
