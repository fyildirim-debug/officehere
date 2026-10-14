// Başlık çubuğundaki kayıt durumu rozetinin kaynağı. Pencere başına tek belge
// olduğundan durum modül düzeyinde tutulur: kaydetme akışları yazar, TitleBar okur.
import { useSyncExternalStore } from "react";
import { getNative, SaveError, type FileRef } from "@officehere/core";

/** Rozetin gösterdiği durum. */
export type SaveState = "dirty" | "saving" | "saved" | "error" | "paused";

export interface SaveStatus {
  /** Son kaydetme girişimi: yok, sürüyor, başarılı ya da başarısız. */
  phase: "idle" | "saving" | "saved" | "error";
  /** Son başarılı kaydetme zamanı (ms). */
  savedAt?: number;
  /** Son hatanın kısa Türkçe nedeni ("Disk dolu"). */
  error?: string;
  /** Son hatanın uzun iletisi (neden + öneri). */
  errorDetail?: string;
  /** Otomatik kaydetme bu belgede duraklatıldıysa nedeni. */
  paused?: string;
  /** "hold": ilk elle kaydı bekliyor; "off": bu dosyada kapalı (doğrulanamadı, yedek alınamadı…). */
  pausedKind?: "hold" | "off";
  /** Belge kimliği: her yeni belge yüklemesinde (reset) artar. */
  gen: number;
  /** Bu belge en son hangi kimlikte elle kaydedildi (otomatik kaydetme beklemesini kaldırır). */
  manualGen: number;
  /** İlk elle kayda kadar otomatik kaydetmeyi bekleten neden (ör. kurtarılan taslak) ve ait olduğu belge kimliği. */
  hold?: string;
  holdGen?: number;
}

/** Kullanıcıya bir kez gösterilecek önemli kayıt uyarısı (ör. doğrulama başarısız). */
export interface SaveAlert {
  title: string;
  message: string;
}

let status: SaveStatus = { phase: "idle", gen: 0, manualGen: -1 };
let retry: (() => void) | null = null;
let alert: SaveAlert | null = null;
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
  saved: (at = Date.now()) => set({ ...status, phase: "saved", savedAt: at, error: undefined, errorDetail: undefined }),
  /** Kaydetme başarısız; neden verilirse rozette gösterilir. */
  failed: (reason?: string, detail?: string) => set({ ...status, phase: "error", error: reason, errorDetail: detail }),
  /** Otomatik kaydetmeyi bu belgede duraklat (null: sürdür). */
  pause: (reason: string | null, kind: "hold" | "off" = "hold") => {
    if ((status.paused ?? null) !== reason || (reason && status.pausedKind !== kind)) set({ ...status, paused: reason ?? undefined, pausedKind: reason ? kind : undefined });
  },
  /**
   * Geçerli belgede ilk ELLE kayda kadar otomatik kaydetme yapılmasın (ör. kurtarılan taslak
   * ya da uyarıyla açılan dosya). Belge yüklendikten (reset) sonra çağrılmalıdır.
   */
  hold: (reason: string) => set({ ...status, hold: reason, holdGen: status.gen }),
  /** Elle Kaydet başarılı: açılış uyarısı nedeniyle bekleyen otomatik kaydetme serbest kalır. */
  noteManualSave: () => {
    if (status.manualGen !== status.gen) set({ ...status, manualGen: status.gen });
  },
  /** Yeni belge açıldı/oluşturuldu: önceki belgenin "Kaydedildi" bilgisi ve duraklatma nedeni taşınmaz. */
  reset: () => set({ phase: "idle", gen: status.gen + 1, manualGen: status.manualGen, hold: status.hold, holdGen: status.holdGen }),
  /** "Kaydedilemedi" rozetine tıklanınca çalışacak normal Kaydet. */
  setRetry(fn: (() => void) | null) {
    retry = fn;
  },
  retry: () => retry?.(),
  /** Önemli uyarı (TitleBar iletişim kutusu olarak gösterir). */
  alert(a: SaveAlert | null) {
    alert = a;
    listeners.forEach((l) => l());
  },
  getAlert: (): SaveAlert | null => alert,
};

/** Kirli bayrağı ve son girişimden rozet durumu (rozet yoksa null). */
export function resolveSaveState(dirty: boolean, s: Pick<SaveStatus, "phase"> & Partial<SaveStatus>): SaveState | null {
  if (s.phase === "saving") return "saving";
  if (s.phase === "error") return "error";
  if (dirty && s.paused) return "paused";
  if (dirty) return "dirty";
  return s.phase === "saved" ? "saved" : null;
}

export function useSaveStatus(): SaveStatus {
  return useSyncExternalStore(saveStatus.subscribe, saveStatus.get, saveStatus.get);
}

export function useSaveAlert(): SaveAlert | null {
  return useSyncExternalStore(saveStatus.subscribe, saveStatus.getAlert, saveStatus.getAlert);
}

/**
 * Elle Kaydet'i rozete yansıtır. `direct`: konum sorulmadan var olan dosyaya
 * yazılacaksa "Kaydediliyor…" gösterilir ve null sonuç hata sayılır; konum
 * soruluyorsa null iptal demektir, rozet değişmez. Yazma hatası (SaveError:
 * "Disk dolu", "Yazma izni yok"…) rozette nedeniyle gösterilir ve null döner;
 * diğer hatalar yeniden fırlatılır.
 */
export async function trackSave<T>(direct: boolean, run: () => Promise<T | null>): Promise<T | null> {
  if (direct) saveStatus.saving();
  try {
    const r = await run();
    if (r) {
      saveStatus.saved();
      saveStatus.noteManualSave();
    } else if (direct) saveStatus.failed();
    return r;
  } catch (err) {
    if (err instanceof SaveError) {
      saveStatus.failed(err.reason, err.message);
      return null;
    }
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
