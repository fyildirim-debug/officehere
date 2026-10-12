import { useCallback, useRef, useState, type ReactNode } from "react";
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
 */
export function RibbonGallery({ items, selectedId, onSelect, itemWidth = 72, itemHeight = 60, columns = 6, footer, expandedColumns }: RibbonGalleryProps) {
  const [row, setRow] = useState(0);
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const close = useCallback(() => setOpen(false), []);
  const rows = Math.max(1, Math.ceil(items.length / columns));
  const visible = items.slice(row * columns, row * columns + columns);

  const tile = (it: GalleryItem, inMenu: boolean) => (
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
