// Mini araç çubuğunun yerleşimi (DOM'dan bağımsız; birim testleri için ayrı).

export interface MiniBarAnchor {
  x: number;
  y: number;
  /** Bağlam menüsüyle birlikte mi açıldı. */
  context: boolean;
  /** Seçimin alt kenarı: üstte yer yoksa çubuk seçimin altına iner. */
  bottom?: number;
}

export interface MiniBarRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/**
 * Seçim kipinde çubuk seçimin üstünde (biraz solda), bağlam kipinde menünün üstünde durur;
 * üstte yer yoksa seçimin ya da menünün altına iner. Yatayda pencereden taşmaz.
 */
export function miniBarPosition(anchor: MiniBarAnchor, size: { w: number; h: number }, view: { w: number; h: number }, menu?: MiniBarRect | null): { left: number; top: number } {
  const clampX = (x: number) => Math.min(Math.max(4, x), view.w - size.w - 4);
  if (anchor.context) {
    if (!menu) return { left: clampX(anchor.x), top: Math.max(4, anchor.y - size.h - 6) };
    let top = menu.top - size.h - 6;
    if (top < 4) top = menu.bottom + 6 + size.h <= view.h - 4 ? menu.bottom + 6 : 4;
    return { left: clampX(menu.left), top };
  }
  let top = anchor.y - size.h - 10;
  if (top < 4) top = anchor.bottom !== undefined ? Math.max(4, Math.min(anchor.bottom + 8, view.h - size.h - 4)) : anchor.y + 28;
  return { left: clampX(anchor.x - 20), top };
}

/** Noktaya en yakın dikdörtgen (en çok `max` px uzaklıkta); bağlam menüsünü bulmak için. */
export function nearestRect(x: number, y: number, rects: MiniBarRect[], max = 8): MiniBarRect | null {
  let best: MiniBarRect | null = null;
  let bestD = max;
  for (const r of rects) {
    const d = Math.hypot(Math.max(r.left - x, 0, x - r.right), Math.max(r.top - y, 0, y - r.bottom));
    if (d <= bestD) {
      best = r;
      bestD = d;
    }
  }
  return best;
}
