// Başlık çubuğundaki kayıt durumu rozetinin kaynağı. Pencere başına tek belge
// olduğundan durum modül düzeyinde tutulur: kaydetme akışları yazar, TitleBar okur.
import { useSyncExternalStore } from "react";
import { getNative, type FileRef } from "@officehere/core";

/** Rozetin gösterdiği durum. */
export type SaveState = "dirty" | "saving" | "saved" | "error";

export interface SaveStatus {
  /** Son kaydetme girişimi: yok, sürüyor, başarılı ya da başarısız. */
  phase: "idle" | "saving" | "saved" | "error";
  /** Son başarılı kaydetme zamanı (ms). */
  savedAt?: number;
}

let status: SaveStatus = { phase: "idle" };
let retry: (() => void) | null = null;
const listeners = new Set<() => void>();

function set(next: SaveStatus) {
  status = next;
  listeners.forEach((l) => l());
}

export const saveStatus = {
  get: (): SaveStatus => status,
  subscribe(cb: () => void): () => void {
    listeners.add(cb);
    return () => listeners.delete(cb);
  },
  saving: () => set({ ...status, phase: "saving" }),
  saved: (at = Date.now()) => set({ phase: "saved", savedAt: at }),
  failed: () => set({ ...status, phase: "error" }),
  /** Yeni belge açıldı/oluşturuldu: önceki belgenin "Kaydedildi" bilgisi taşınmaz. */
  reset: () => {
    if (status.phase !== "idle") set({ phase: "idle" });
  },
  /** "Kaydedilemedi" rozetine tıklanınca çalışacak normal Kaydet. */
  setRetry(fn: (() => void) | null) {
    retry = fn;
  },
  retry: () => retry?.(),
};

/** Kirli bayrağı ve son girişimden rozet durumu (rozet yoksa null). */
export function resolveSaveState(dirty: boolean, s: SaveStatus): SaveState | null {
  if (s.phase === "saving") return "saving";
  if (s.phase === "error") return "error";
  if (dirty) return "dirty";
  return s.phase === "saved" ? "saved" : null;
}

export function useSaveStatus(): SaveStatus {
  return useSyncExternalStore(saveStatus.subscribe, saveStatus.get, saveStatus.get);
}

/**
 * Elle Kaydet'i rozete yansıtır. `direct`: konum sorulmadan var olan dosyaya
 * yazılacaksa "Kaydediliyor…" gösterilir ve null sonuç hata sayılır; konum
 * soruluyorsa null iptal demektir, rozet değişmez.
 */
export async function trackSave<T>(direct: boolean, run: () => Promise<T | null>): Promise<T | null> {
  if (direct) saveStatus.saving();
  try {
    const r = await run();
    if (r) saveStatus.saved();
    else if (direct) saveStatus.failed();
    return r;
  } catch (err) {
    saveStatus.failed();
    throw err;
  }
}

/** Kaydet konum sormadan doğrudan bu dosyaya mı yazacak (masaüstünde yol, tarayıcıda dosya tanıtıcısı)? */
export function isDirectSave(ref: FileRef | null | undefined): boolean {
  if (!ref) return false;
  return ref.path ? !!getNative() : !!ref.handle;
}

/** HH:MM (Türkçe yerel ayar). */
export function formatSavedTime(at: number): string {
  return new Date(at).toLocaleTimeString("tr-TR", { hour: "2-digit", minute: "2-digit" });
}
