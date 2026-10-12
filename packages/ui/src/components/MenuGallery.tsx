import type { ReactNode } from "react";
import { useMenuClose } from "./Menu";
import { tipProps } from "./Tooltip";

export interface MenuGalleryItem {
  id: string;
  /** Erişilebilir ad (ekran okuyucu, harfle atlama). */
  label: string;
  /** İpucu başlığı (yoksa label) ve açıklaması. */
  tip?: string;
  description?: string;
  preview: ReactNode;
}

/**
 * Açılır menü içinde önizlemeli seçim ızgarası (ör. Konum galerisi). Kutucuklar menuitemradio rolündedir;
 * ok tuşlarıyla ızgarada gezinme, Enter/Boşluk ile seçim ve Esc menünün Popover klavye desteğinden gelir.
 * Seçim menüyü kapatır.
 */
export function MenuGallery({
  items,
  selectedId,
  onSelect,
  columns = 3,
  itemWidth = 56,
  itemHeight = 64,
  label,
}: {
  items: MenuGalleryItem[];
  selectedId?: string | null;
  onSelect: (id: string) => void;
  columns?: number;
  itemWidth?: number;
  itemHeight?: number;
  label?: string;
}) {
  const close = useMenuClose();
  return (
    <div className="oh-menu-gallery" role="group" aria-label={label} style={{ gridTemplateColumns: `repeat(${columns}, ${itemWidth}px)` }}>
      {items.map((it) => {
        const on = it.id === selectedId;
        return (
          <button
            key={it.id}
            type="button"
            role="menuitemradio"
            aria-checked={on}
            aria-label={it.label}
            tabIndex={-1}
            className={`oh-gallery-item${on ? " is-selected" : ""}`}
            style={{ width: itemWidth, height: itemHeight }}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              onSelect(it.id);
              close();
            }}
            {...tipProps(it.tip ?? it.label, undefined, it.description)}
          >
            <span className="oh-gallery-preview">{it.preview}</span>
          </button>
        );
      })}
    </div>
  );
}
