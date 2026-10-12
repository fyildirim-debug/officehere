// Otomatik kaydetmenin React'ten bağımsız çekirdeği: seçenek, biçim denetimi,
// oturumluk yedek, boşta kalma zamanlayıcısı ve tek bir kaydetme turu.
import { extensionOf } from "@officehere/core";

/** Seçenekler'deki "Otomatik kaydet" (üç uygulamada ortak). */
export const AUTO_SAVE_KEY = "officehere.autoSave";
/** Aynı penceredeki dinleyicilere seçenek değişikliğini duyuran olay. */
export const AUTO_SAVE_EVENT = "officehere:autosave-pref";
/** Son değişiklikten sonra beklenen boşluk. */
export const AUTO_SAVE_IDLE_MS = 2000;
/** Sürekli düzenlemede en geç bu sürede bir kaydedilir. */
export const AUTO_SAVE_MAX_WAIT_MS = 30_000;
/** Biçim kaybı olmadan gidiş-dönüş yazılabilen kendi biçimlerimiz. */
export const AUTO_SAVE_FORMATS = ["docx", "xlsx", "pptx"] as const;

type StorageLike = Pick<Storage, "getItem" | "setItem">;

function defaultStorage(): StorageLike | null {
  try {
    return typeof localStorage === "undefined" ? null : localStorage;
  } catch {
    return null;
  }
}

/** Varsayılan açık; yalnız açıkça kapatılmışsa kapalı. */
export function readAutoSavePref(storage: StorageLike | null = defaultStorage()): boolean {
  try {
    const v = storage?.getItem(AUTO_SAVE_KEY);
    return v !== "0" && v !== "false";
  } catch {
    return true;
  }
}

export function writeAutoSavePref(on: boolean, storage: StorageLike | null = defaultStorage()): void {
  try {
    storage?.setItem(AUTO_SAVE_KEY, on ? "1" : "0");
  } catch {
    // Yerel depo kapalıysa seçenek bu oturumda kalıcı olmaz.
  }
  if (typeof window !== "undefined") window.dispatchEvent(new Event(AUTO_SAVE_EVENT));
}

/** Dosya otomatik kaydedilebilir mi: diskte yolu var ve uzantısı kendi biçimimiz. */
export function canAutoSavePath(path: string | undefined | null, formats: readonly string[] = AUTO_SAVE_FORMATS): path is string {
  if (!path) return false;
  return formats.includes(extensionOf(path));
}

/** Yedek dosyanın yolu: "rapor.docx" → "rapor.docx.bak" (özgün uzantı korunur, geri adlandırmak kolay). */
export function backupPathFor(path: string): string {
  return `${path}.bak`;
}

/** Disk erişimi (masaüstünde köprü; testlerde sahte). */
export interface AutoSaveIo {
  read(path: string): Promise<ArrayBuffer | null>;
  write(path: string, data: ArrayBuffer): Promise<boolean>;
}

/** Bu pencerede yedeği alınmış (ya da alınamamış) dosyalar. */
const backups = new Map<string, Promise<boolean>>();

/**
 * Bir dosyanın özgün hâlinin yedeğini oturumda bir kez alır. Okunamaz ya da
 * yazılamazsa false döner ve sonuç akılda tutulur: o dosya bu oturumda otomatik
 * kaydedilmez.
 */
export function ensureBackup(path: string, io: AutoSaveIo): Promise<boolean> {
  let p = backups.get(path);
  if (!p) {
    p = (async () => {
      try {
        const data = await io.read(path);
        if (!data) return false;
        return await io.write(backupPathFor(path), data);
      } catch {
        return false;
      }
    })();
    backups.set(path, p);
  }
  return p;
}

/** Yalnız testler için: oturum yedek kaydını sıfırlar. */
export function resetBackupsForTest(): void {
  backups.clear();
}

export interface IdleScheduler {
  /** Bir değişiklik oldu: boşluk süresini yeniden başlat (en geç üst sınırda çalışır). */
  touch(): void;
  cancel(): void;
  readonly pending: boolean;
}

/** Boşta kalma zamanlayıcısı: son dokunuştan idleMs sonra, sürekli dokunuşta en geç maxWaitMs içinde çalışır. */
export function createIdleScheduler(run: () => void, idleMs = AUTO_SAVE_IDLE_MS, maxWaitMs = AUTO_SAVE_MAX_WAIT_MS, now: () => number = Date.now): IdleScheduler {
  let timer: ReturnType<typeof setTimeout> | undefined;
  let first = 0;
  const fire = () => {
    timer = undefined;
    first = 0;
    run();
  };
  return {
    touch() {
      const t = now();
      if (!first) first = t;
      if (timer !== undefined) clearTimeout(timer);
      timer = setTimeout(fire, Math.max(0, Math.min(idleMs, first + maxWaitMs - t)));
    },
    cancel() {
      if (timer !== undefined) clearTimeout(timer);
      timer = undefined;
      first = 0;
    },
    get pending() {
      return timer !== undefined;
    },
  };
}

/** Uygulamanın kendi biçiminde ürettiği içerik ve o andaki belge sürümünün hâlâ geçerli olup olmadığı. */
export interface PreparedSave {
  data: ArrayBuffer | Uint8Array;
  /** Yazma bitince belge hâlâ bu içerikle aynı mı (değilse kirli kalır). */
  isCurrent: () => boolean;
}

export interface AutoSaveTurn {
  path: string | undefined;
  formats?: readonly string[];
  enabled: boolean;
  isDirty: () => boolean;
  prepare: () => Promise<PreparedSave>;
  markClean: () => void;
}

export interface AutoSaveReport {
  saving(): void;
  saved(): void;
  failed(): void;
}

export type AutoSaveResult = "skipped" | "no-backup" | "saved" | "changed" | "failed";

function toArrayBuffer(data: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (data instanceof ArrayBuffer) return data;
  return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;
}

/**
 * Tek bir otomatik kaydetme turu: kapalıysa, belge temizse ya da biçim uygun
 * değilse atlar; yedek alınamazsa yazmaz (rozet "Kaydedilmedi" kalır); yazma
 * başarılıysa ve belge bu arada değişmediyse kirli durumu temizler.
 */
export async function runAutoSave(turn: AutoSaveTurn, io: AutoSaveIo, report: AutoSaveReport): Promise<AutoSaveResult> {
  const { path } = turn;
  if (!turn.enabled || !turn.isDirty() || !canAutoSavePath(path, turn.formats)) return "skipped";
  if (!(await ensureBackup(path, io))) return "no-backup";
  if (!turn.isDirty()) return "skipped";
  report.saving();
  try {
    const prepared = await turn.prepare();
    const ok = await io.write(path, toArrayBuffer(prepared.data));
    if (!ok) {
      report.failed();
      return "failed";
    }
    report.saved();
    if (!prepared.isCurrent()) return "changed";
    turn.markClean();
    return "saved";
  } catch {
    report.failed();
    return "failed";
  }
}
