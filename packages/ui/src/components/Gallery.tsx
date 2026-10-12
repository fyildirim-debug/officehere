import { useCallback, useEffect, useRef, useState, type KeyboardEvent as ReactKeyboardEvent, type ReactNode } from "react";
import { Popover } from "./Popover";
import { MenuProvider } from "./Menu";
import { tipProps } from "./Tooltip";

export interface GalleryItem {
  id: string;
  label: string;
  /** Kutucuk içinde çizilecek önizleme. */
  preview: ReactNode;
}

export interface RibbonGalleryProps {
  items: GalleryItem[];
  selectedId?: string;
  onSelect: (id: string) => void;
  /** Kutucuk boyutu (px). */
  itemWidth?: number;
  itemHeight?: number;
  /** Şeritte görünen sütun sayısı. */
  columns?: number;
  /** Açılmış galerinin altındaki ek menü (ör. "Stil Oluştur"). */
  footer?: ReactNode;
  /** Açılmış galerideki sütun sayısı. */
  expandedColumns?: number;
}

/**
 * Şerit içi galeri (Stiller, Şekil Stilleri, Tasarım temaları). Sağdaki
 * oklarla satır kaydırılır, alttaki düğme galeriyi açılır menüde genişletir.
 * Klavye: galeri tek Tab durağıdır; ←/→ öğeler arasında (satır gerekirse kayar), ↑/↓ satır
 * değiştirir, Home/End, Alt+↓ genişletir. Açılmış galeride oklar ızgarada gezer (Popover).
 */
export function RibbonGallery({ items, selectedId, onSelect, itemWidth = 72, itemHeight = 60, columns = 6, footer, expandedColumns }: RibbonGalleryProps) {
  const [row, setRow] = useState(0);
  const [open, setOpen] = useState(false);
  // Klavye odağındaki öğe (dolaşan tabindex); yoksa seçili ya da görünen ilk öğe.
  const [focusId, setFocusId] = useState<string | null>(null);
  const ref = useRef<HTMLDivElement>(null);
  const itemsRef = useRef<HTMLDivElement>(null);
  const pendingFocus = useRef(false);
  const close = useCallback(() => setOpen(false), []);
  const rows = Math.max(1, Math.ceil(items.length / columns));
  const visible = items.slice(row * columns, row * columns + columns);
  const rovingId = visible.find((it) => it.id === focusId)?.id ?? visible.find((it) => it.id === selectedId)?.id ?? visible[0]?.id;

  useEffect(() => {
    if (!pendingFocus.current) return;
    pendingFocus.current = false;
    itemsRef.current?.querySelector<HTMLElement>('[tabindex="0"]')?.focus({ preventScroll: true });
  });

  const onKey = (e: ReactKeyboardEvent) => {
    const i = items.findIndex((it) => it.id === rovingId);
    if (i < 0) return;
    let next = -1;
    if (e.key === "ArrowDown" && e.altKey) {
      e.preventDefault();
      setOpen(true);
      return;
    }
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    if (e.key === "ArrowRight") next = Math.min(items.length - 1, i + 1);
    else if (e.key === "ArrowLeft") next = Math.max(0, i - 1);
    else if (e.key === "ArrowDown") next = Math.min(items.length - 1, i + columns);
    else if (e.key === "ArrowUp") next = i - columns >= 0 ? i - columns : i;
    else if (e.key === "Home") next = 0;
    else if (e.key === "End") next = items.length - 1;
    else return;
    e.preventDefault();
    const it = items[next];
    if (!it || next === i) return;
    setFocusId(it.id);
    setRow(Math.floor(next / columns));
    pendingFocus.current = true;
  };

  const tile = (it: GalleryItem, inMenu: boolean) => (
    <button
      key={it.id}
      type="button"
      className={`oh-gallery-item${it.id === selectedId ? " is-selected" : ""}`}
      style={{ width: itemWidth, height: itemHeight }}
      aria-label={it.label}
      role={inMenu ? "menuitemradio" : undefined}
      aria-checked={inMenu ? it.id === selectedId : undefined}
      aria-pressed={inMenu ? undefined : it.id === selectedId}
      tabIndex={inMenu ? -1 : it.id === rovingId ? 0 : -1}
      onMouseDown={(e) => e.preventDefault()}
      onFocus={inMenu ? undefined : () => setFocusId(it.id)}
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
      <div ref={itemsRef} className="oh-gallery-items" role="group" style={{ width: columns * (itemWidth + 2) + 2, height: itemHeight + 4 }} onKeyDown={onKey}>
        {visible.map((it) => tile(it, false))}
      </div>
      <div className="oh-gallery-scroll">
        <button type="button" tabIndex={-1} disabled={row === 0} onMouseDown={(e) => e.preventDefault()} onClick={() => setRow((r) => Math.max(0, r - 1))} aria-label="Yukarı">
          <svg width="8" height="5" viewBox="0 0 8 5" aria-hidden="true">
            <path d="M0 5l4-5 4 5z" fill="currentColor" />
          </svg>
        </button>
        <button type="button" tabIndex={-1} disabled={row >= rows - 1} onMouseDown={(e) => e.preventDefault()} onClick={() => setRow((r) => Math.min(rows - 1, r + 1))} aria-label="Aşağı">
          <svg width="8" height="5" viewBox="0 0 8 5" aria-hidden="true">
            <path d="M0 0l4 5 4-5z" fill="currentColor" />
          </svg>
        </button>
        <button
          type="button"
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => setOpen(true)}
          aria-label="Diğer"
          aria-haspopup="menu"
          aria-expanded={open}
          onKeyDown={(e) => {
            if (e.key === "ArrowDown" && !open) (e.preventDefault(), setOpen(true));
          }}
          {...tipProps("Diğer")}
        >
          <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden="true">
            <path d="M0 1h8" stroke="currentColor" />
            <path d="M0 3l4 5 4-5z" fill="currentColor" />
          </svg>
        </button>
      </div>
      <Popover open={open} onClose={close} anchor={ref}>
        <MenuProvider onClose={close}>
          <div className="oh-menu oh-gallery-menu" role="menu" tabIndex={-1}>
            <div className="oh-gallery-grid" style={{ gridTemplateColumns: `repeat(${expandedColumns ?? columns}, ${itemWidth}px)` }}>
              {items.map((it) => tile(it, true))}
            </div>
            {footer}
          </div>
        </MenuProvider>
      </Popover>
    </div>
  );
}
