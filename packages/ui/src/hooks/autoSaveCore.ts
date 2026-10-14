// Otomatik kaydetmenin React'ten bağımsız çekirdeği: seçenek, biçim denetimi,
// boşta kalma zamanlayıcısı, aralığın seyreltilmesi ve tek bir kaydetme turu.
//
// Güvenlik kuralları:
// - Yazma ana sürecin güvenli yolundan geçer (geçici dosya + fsync + taşıma, yol başına kuyruk).
// - İlk otomatik yazmadan önce özgün dosya "<ad>.bak" olarak yedeklenir; .bak yalnız
//   yoksa oluşturulur, oturumlar arasında ezilmez (ana süreç yapar). Yedek alınamazsa yazılmaz.
// - Dosya uyarıyla/eksik açıldıysa ilk ELLE kayda kadar otomatik kaydetme yapılmaz.
// - Yazılan dosya uygulamanın kendi okuyucusuyla yeniden açılabiliyor mu denetlenir; açılamazsa
//   önceki içerik geri yazılır, otomatik kaydetme o dosyada kapanır ve kullanıcı uyarılır.
import { extensionOf, saveErrorReason, type NativeWriteOptions, type WriteOutcome } from "@officehere/core";

/** Seçenekler'deki "Otomatik kaydet" (üç uygulamada ortak). */
export const AUTO_SAVE_KEY = "officehere.autoSave";
/** Aynı penceredeki dinleyicilere seçenek değişikliğini duyuran olay. */
export const AUTO_SAVE_EVENT = "officehere:autosave-pref";
/** Son değişiklikten sonra beklenen boşluk. */
export const AUTO_SAVE_IDLE_MS = 2000;
/** Sürekli düzenlemede en geç bu sürede bir kaydedilir. */
export const AUTO_SAVE_MAX_WAIT_MS = 30_000;
/** Büyük dosyada aralık son kaydın (üretme + yazma + doğrulama) süresinin bu katına çıkar. */
export const AUTO_SAVE_COST_FACTOR = 10;
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

/** Disk erişimi (masaüstünde ana süreç köprüsü; testlerde sahte). */
export interface AutoSaveIo {
  /** Güvenli yazma; `sidecar`: ".bak" yoksa önce alınır, alınamazsa yazılmaz (EBACKUP). */
  write(path: string, data: ArrayBuffer, options: NativeWriteOptions): Promise<WriteOutcome>;
  read(path: string): Promise<ArrayBuffer | null>;
  /** keepPrevious ile yazılan dosyada son yazmadan önceki içeriği geri yazar. */
  revert(path: string): Promise<WriteOutcome>;
}

/** Bu pencerede otomatik kaydetmesi kapatılmış dosyalar (doğrulama başarısız) ve nedeni. */
const disabled = new Map<string, string>();

/** Dosyada otomatik kaydetme kapatıldıysa nedeni. */
export function autoSaveDisabledReason(path: string | undefined): string | undefined {
  return path ? disabled.get(path) : undefined;
}

/** Bu pencerede dosyanın otomatik kaydedilmesini kapatır (ör. dosya başka pencerede de açık). */
export function disableAutoSave(path: string, reason: string): void {
  disabled.set(path, reason);
}

/** Yalnız testler için: oturum kayıtlarını sıfırlar. */
export function resetAutoSaveForTest(): void {
  disabled.clear();
}

export interface IdleScheduler {
  /** Bir değişiklik oldu: boşluk süresini yeniden başlat (en geç üst sınırda çalışır). */
  touch(): void;
  cancel(): void;
  readonly pending: boolean;
}

type Ms = number | (() => number);
const ms = (v: Ms) => (typeof v === "function" ? v() : v);

/**
 * Boşta kalma zamanlayıcısı: son dokunuştan idleMs sonra, sürekli dokunuşta en geç
 * maxWaitMs içinde çalışır. Süreler işlev olarak verilirse her dokunuşta yeniden okunur.
 */
export function createIdleScheduler(run: () => void, idleMs: Ms = AUTO_SAVE_IDLE_MS, maxWaitMs: Ms = AUTO_SAVE_MAX_WAIT_MS, now: () => number = Date.now): IdleScheduler {
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
      timer = setTimeout(fire, Math.max(0, Math.min(ms(idleMs), first + ms(maxWaitMs) - t)));
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

/**
 * Aralığın seyreltilmesi: son kayıt (üretme + yazma + doğrulama) uzun sürdüyse bekleme o
 * sürenin 10 katına çıkar; böylece büyük dosyada otomatik kaydetme zamanın en çok ~%10'unu alır.
 */
export function autoSaveDelays(lastCostMs: number): { idleMs: number; maxWaitMs: number } {
  const scaled = Math.max(0, lastCostMs) * AUTO_SAVE_COST_FACTOR;
  return { idleMs: Math.max(AUTO_SAVE_IDLE_MS, scaled), maxWaitMs: Math.max(AUTO_SAVE_MAX_WAIT_MS, scaled) };
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
  /** Otomatik kaydetme bekletiliyorsa nedeni (ör. dosya uyarıyla açıldı; ilk elle kayda kadar). */
  hold?: string | null;
  /**
   * Diske yazılan baytlar uygulamanın kendi okuyucusuyla yeniden açılabiliyor mu.
   * false dönerse ya da hata fırlatırsa önceki içerik geri yazılır.
   */
  verify?: (data: ArrayBuffer) => Promise<boolean>;
  /** Doğrulamadan önce beklenen boşta kalma anı (büyük dosyada arayüzü kilitlememek için). */
  idle?: () => Promise<void>;
  now?: () => number;
}

export interface AutoSaveReport {
  saving(): void;
  saved(): void;
  /** Yazma başarısız; kısa neden ("Disk dolu") ve uzun ileti. */
  failed(reason?: string, detail?: string): void;
  /** Otomatik kaydetme duraklatıldı (null: sürüyor); "hold": ilk elle kaydı bekliyor, "off": bu dosyada kapalı. */
  paused?(reason: string | null, kind?: "hold" | "off"): void;
  /** Kullanıcı uyarılmalı (doğrulama başarısız, önceki sürüm geri yüklendi). */
  alert?(title: string, message: string): void;
}

export type AutoSaveResult = "skipped" | "held" | "disabled" | "no-backup" | "saved" | "changed" | "failed" | "verify-failed";

function toArrayBuffer(data: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (data instanceof ArrayBuffer) return data;
  return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;
}

/** Bir turun ayrıntısı: sonuç ve harcanan süre (aralığı seyreltmek için). */
export interface AutoSaveRun {
  result: AutoSaveResult;
  costMs: number;
}

/**
 * Tek bir otomatik kaydetme turu: kapalıysa, belge temizse ya da biçim uygun
 * değilse atlar; bekletiliyorsa (uyarıyla açılmış dosya) yazmaz; yedek alınamazsa
 * yazmaz ve duraklar; yazma başarılıysa doğrular, doğrulanamazsa geri alır ve
 * o dosyada otomatik kaydetmeyi kapatır; belge bu arada değişmediyse kirli
 * durumu temizler.
 */
export async function runAutoSaveTurn(turn: AutoSaveTurn, io: AutoSaveIo, report: AutoSaveReport): Promise<AutoSaveRun> {
  const now = turn.now ?? Date.now;
  const t0 = now();
  const done = (result: AutoSaveResult): AutoSaveRun => ({ result, costMs: now() - t0 });
  const { path } = turn;
  if (!turn.enabled || !turn.isDirty() || !canAutoSavePath(path, turn.formats)) return done("skipped");
  const off = disabled.get(path);
  if (off) {
    report.paused?.(off, "off");
    return done("disabled");
  }
  if (turn.hold) {
    report.paused?.(turn.hold, "hold");
    return done("held");
  }
  report.paused?.(null);
  report.saving();
  let prepared: PreparedSave;
  let out: WriteOutcome;
  try {
    prepared = await turn.prepare();
    out = await io.write(path, toArrayBuffer(prepared.data), { sidecar: true, keepPrevious: !!turn.verify });
  } catch (err) {
    report.failed(undefined, (err as Error)?.message);
    return done("failed");
  }
  if (!out.ok) {
    if (out.code === "EBACKUP") {
      // Özgün dosya korunamadı: üzerine yazılmadı. Kullanıcı Kaydet ile kendisi karar verir.
      const reason = "Özgün dosyanın yedeği alınamadı";
      disabled.set(path, reason);
      report.failed(saveErrorReason("EBACKUP"), out.message);
      report.paused?.(reason, "off");
      return done("no-backup");
    }
    report.failed(out.reason ?? saveErrorReason(out.code), out.message);
    return done("failed");
  }
  if (turn.verify) {
    let good = false;
    try {
      await turn.idle?.();
      const bytes = await io.read(path);
      good = !!bytes && (await turn.verify(bytes)) !== false;
    } catch {
      good = false;
    }
    if (!good) {
      const reverted = await io.revert(path).catch((): WriteOutcome => ({ ok: false, code: "EUNKNOWN" }));
      const reason = "Kaydedilen dosya doğrulanamadı";
      disabled.set(path, reason);
      report.failed(reason, reverted.ok ? "Önceki içerik geri yüklendi." : "Önceki içerik geri yüklenemedi.");
      report.paused?.(reason, "off");
      report.alert?.(
        "Otomatik kaydetme kapatıldı",
        reverted.ok
          ? "Otomatik kaydedilen dosya yeniden açılabilir olduğu doğrulanamadı. Dosyanın önceki içeriği geri yüklendi ve bu dosya için otomatik kaydetme kapatıldı. Çalışmanız açık ve kaydedilmemiş durumda; Farklı Kaydet ile yeni bir dosyaya kaydetmeniz önerilir."
          : "Otomatik kaydedilen dosya yeniden açılabilir olduğu doğrulanamadı ve önceki içerik geri yüklenemedi. Özgün dosya \".bak\" uzantılı yedekte duruyor. Çalışmanız açık; Farklı Kaydet ile yeni bir dosyaya kaydedin.",
      );
      return done("verify-failed");
    }
  }
  report.saved();
  if (!prepared.isCurrent()) return done("changed");
  turn.markClean();
  return done("saved");
}

/** Geriye uyumluluk: yalnız sonucu döner. */
export async function runAutoSave(turn: AutoSaveTurn, io: AutoSaveIo, report: AutoSaveReport): Promise<AutoSaveResult> {
  return (await runAutoSaveTurn(turn, io, report)).result;
}
