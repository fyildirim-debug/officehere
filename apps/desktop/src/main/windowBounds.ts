// Pencere konumu ve boyutu hesapları (Electron'dan bağımsız; birim testli).

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface Size {
  width: number;
  height: number;
}

/** pencere.json'da uygulama türü başına saklanan durum. */
export interface SavedWindow {
  x?: number;
  y?: number;
  width: number;
  height: number;
  maximized?: boolean;
}

export interface WindowPlan {
  /** Verilmezse Electron pencereyi birincil ekranda ortalar. */
  x?: number;
  y?: number;
  width: number;
  height: number;
  maximized: boolean;
}

/** Aynı türden yeni pencerenin bir öncekine göre kayması. */
export const CASCADE_STEP = 24;

const finite = (n: unknown): n is number => typeof n === "number" && Number.isFinite(n);

/** Bozuk ya da eksik kaydı ayıklar. */
export function sanitizeSaved(raw: unknown): SavedWindow | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const r = raw as Record<string, unknown>;
  if (!finite(r.width) || !finite(r.height) || r.width < 100 || r.height < 100) return undefined;
  const out: SavedWindow = { width: Math.round(r.width), height: Math.round(r.height), maximized: r.maximized === true };
  if (finite(r.x) && finite(r.y)) {
    out.x = Math.round(r.x);
    out.y = Math.round(r.y);
  }
  return out;
}

function overlap(a: Rect, b: Rect): { w: number; h: number } {
  return {
    w: Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x),
    h: Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y),
  };
}

/** Pencerenin başlık çubuğu (üst 32 px) bir ekranın çalışma alanında tutulabilecek kadar görünüyor mu. */
export function isReachable(rect: Rect, areas: readonly Rect[]): boolean {
  const strip = { x: rect.x, y: rect.y, width: rect.width, height: 32 };
  return areas.some((a) => {
    const o = overlap(strip, a);
    return o.w >= Math.min(80, rect.width) && o.h >= 16;
  });
}

/** Dikdörtgenle en çok kesişen çalışma alanı. */
export function bestArea(rect: Rect, areas: readonly Rect[]): Rect | undefined {
  let best: Rect | undefined;
  let bestArea = 0;
  for (const a of areas) {
    const o = overlap(rect, a);
    const area = Math.max(0, o.w) * Math.max(0, o.h);
    if (area > bestArea) {
      best = a;
      bestArea = area;
    }
  }
  return best;
}

export function centerIn(size: Size, area: Rect): Rect {
  const width = Math.min(size.width, area.width);
  const height = Math.min(size.height, area.height);
  return { x: Math.round(area.x + (area.width - width) / 2), y: Math.round(area.y + (area.height - height) / 2), width, height };
}

/** Bir öncekinin CASCADE_STEP px sağ altı; çalışma alanından taşacaksa alanın sol üst köşesine döner. */
export function cascadeFrom(prev: Rect, area: Rect, step = CASCADE_STEP): Rect {
  const next = { x: prev.x + step, y: prev.y + step, width: prev.width, height: prev.height };
  // Alandan büyük pencerede yalnız başlık çubuğunun erişilebilirliğine bakılır.
  const overflowX = next.width <= area.width && next.x + next.width > area.x + area.width;
  const overflowY = next.height <= area.height && next.y + next.height > area.y + area.height;
  if (overflowX || overflowY || !isReachable(next, [area])) return { ...next, x: area.x, y: area.y };
  return next;
}

export interface PlanInput {
  defaults: Size;
  /** Bu tür için saklanan son durum. */
  saved?: SavedWindow;
  /** Aynı türden açık son pencere (varsa yeni pencere ondan kaydırılır). */
  previous?: Rect & { maximized?: boolean };
  /** Tüm ekranların çalışma alanları. */
  areas: readonly Rect[];
  primary: Rect;
}

/** Yeni pencerenin konumu, boyutu ve büyütülmüş olup olmadığı. */
export function planWindow({ defaults, saved, previous, areas, primary }: PlanInput): WindowPlan {
  if (previous) {
    const area = bestArea(previous, areas) ?? primary;
    return { ...cascadeFrom(previous, area), maximized: previous.maximized === true };
  }
  if (!saved) return { width: defaults.width, height: defaults.height, maximized: false };
  const maximized = saved.maximized === true;
  if (saved.x !== undefined && saved.y !== undefined) {
    const rect = { x: saved.x, y: saved.y, width: saved.width, height: saved.height };
    if (isReachable(rect, areas)) {
      const area = bestArea(rect, areas) ?? primary;
      // Ekran küçülmüşse pencere alana sığdırılır.
      return { ...rect, width: Math.min(rect.width, area.width), height: Math.min(rect.height, area.height), maximized };
    }
  }
  // Kaydedilen konum hiçbir ekrana düşmüyor (ekran çıkarılmış olabilir): birincil ekranda ortala.
  return { ...centerIn(saved, primary), maximized };
}
