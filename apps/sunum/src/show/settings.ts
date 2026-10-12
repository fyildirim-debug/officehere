/**
 * Gösteri ayarları (Gösteriyi Ayarla) ve gösterilecek slaytların sırası:
 * tüm slaytlar, bir aralık ya da özel slayt gösterisi. Gizli slaytlar
 * sıradan çıkarılır (yalnız doğrudan atlanarak gösterilir).
 */
import type { CustomShow, Presentation } from "../model/types";

export type ShowType = "speaker" | "browsed" | "kiosk";

export type ShowRange = { kind: "all" } | { kind: "range"; from: number; to: number } | { kind: "custom"; id: string };

export interface ShowSettings {
  /** Sunucu tarafından sunulan (tam ekran), birey tarafından göz atılan (pencere), bilgi noktası (tam ekran, döngü). */
  type: ShowType;
  loop: boolean;
  noNarration: boolean;
  noAnimations: boolean;
  /** "#RRGGBB" */
  penColor: string;
  laserColor: string;
  range: ShowRange;
  /** El ile ya da varsa zamanlamaları kullanarak. */
  advance: "manual" | "timings";
  /** Sunucu Görünümünü kullan (izleyici ayrı pencerede). */
  presenter: boolean;
}

export const DEFAULT_SHOW_SETTINGS: ShowSettings = {
  type: "speaker",
  loop: false,
  noNarration: false,
  noAnimations: false,
  penColor: "#E81123",
  laserColor: "#FF0000",
  range: { kind: "all" },
  advance: "timings",
  presenter: false,
};

/** Gösteri motoruna giden çalışma ayarları. */
export interface ShowRunOptions {
  loop?: boolean;
  noAnimations?: boolean;
  noNarration?: boolean;
  kiosk?: boolean;
  /** Otomatik ilerleme zamanlamaları kullanılmasın. */
  manual?: boolean;
  penColor?: string;
  laserColor?: string;
  /** Gösterilecek slayt dizinleri (sırayla). */
  order?: number[];
  /** Zamanlama Provası: slayt sürelerini kaydet. */
  rehearse?: boolean;
}

/** Ayarlara göre gösterilecek slayt dizinleri. */
export function showOrder(pres: Presentation, range: ShowRange): number[] {
  const n = pres.slides.length;
  let idx: number[];
  if (range.kind === "range") {
    const from = Math.max(1, Math.min(n, range.from));
    const to = Math.max(from, Math.min(n, range.to));
    idx = [];
    for (let i = from - 1; i < to; i++) idx.push(i);
  } else if (range.kind === "custom") {
    const cs = pres.customShows?.find((c) => c.id === range.id);
    const pos = new Map(pres.slides.map((s, i) => [s.id, i]));
    idx = (cs?.slideIds ?? []).map((id) => pos.get(id)).filter((i): i is number => i !== undefined);
  } else idx = pres.slides.map((_, i) => i);
  return idx.filter((i) => !pres.slides[i]?.hidden);
}

/** Sıradaki (dir=1) ya da önceki (dir=-1) slayt; yoksa -1 / order dışı. */
export function stepInOrder(order: number[], current: number, dir: 1 | -1): number | null {
  if (!order.length) return null;
  const pos = order.indexOf(current);
  if (pos >= 0) {
    const np = pos + dir;
    return np >= 0 && np < order.length ? order[np] : null;
  }
  // Sıra dışı (ör. gizli slayta atlanmış): dizine göre en yakın komşu.
  if (dir > 0) return order.find((i) => i > current) ?? null;
  for (let k = order.length - 1; k >= 0; k--) if (order[k] < current) return order[k];
  return null;
}

/** Ayarlardan gösteri motoru seçenekleri. */
export function runOptions(pres: Presentation, s: ShowSettings): ShowRunOptions {
  const kiosk = s.type === "kiosk";
  return {
    loop: s.loop || kiosk,
    noAnimations: s.noAnimations,
    noNarration: s.noNarration,
    kiosk,
    manual: s.advance === "manual",
    penColor: s.penColor,
    laserColor: s.laserColor,
    order: showOrder(pres, s.range),
  };
}

/** Özel gösteri kopyası için benzersiz ad ("Ad (2)"). */
export function uniqueShowName(list: CustomShow[], base: string): string {
  const names = new Set(list.map((c) => c.name));
  if (!names.has(base)) return base;
  for (let i = 2; ; i++) if (!names.has(`${base} (${i})`)) return `${base} (${i})`;
}
