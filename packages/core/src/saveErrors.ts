// Kaydetme hataları: masaüstü ana süreci yazma hatasının kodunu (ENOSPC, EACCES…)
// ve Türkçe nedenini iletir; sayfalar bunu rozette ve iletilerde gösterir.
// Ana süreçteki karşılığı: apps/desktop/src/main/errors.ts.

/** Ana sürecin yazma sonucu. */
export type WriteOutcome = { ok: true } | { ok: false; code: string; reason?: string; message?: string };

const REASONS: Record<string, string> = {
  ENOSPC: "Disk dolu",
  EDQUOT: "Disk kotası doldu",
  QuotaExceededError: "Disk dolu",
  EACCES: "Yazma izni yok",
  EPERM: "Yazma izni yok",
  NotAllowedError: "Yazma izni yok",
  NoModificationAllowedError: "Dosya başka bir programda açık",
  EBUSY: "Dosya başka bir programda açık",
  ETXTBSY: "Dosya başka bir programda açık",
  EROFS: "Disk salt okunur",
  ENOENT: "Klasör bulunamadı",
  ENOTDIR: "Klasör bulunamadı",
  NotFoundError: "Dosya bulunamadı",
  EISDIR: "Bu adda bir klasör var",
  ENAMETOOLONG: "Dosya adı ya da yolu çok uzun",
  EIO: "Disk okuma/yazma hatası",
  EMFILE: "Çok fazla dosya açık",
  ENFILE: "Çok fazla dosya açık",
  EFBIG: "Dosya bu disk için çok büyük",
  EBACKUP: "Özgün dosyanın yedeği alınamadı",
  EVERIFY: "Kaydedilen dosya doğrulanamadı",
};

/** Hata kodunun kısa Türkçe nedeni ("Disk dolu"). */
export function saveErrorReason(code: string | undefined): string {
  return (code && REASONS[code]) || "Dosya yazılamadı";
}

/** Kaydetme hatası: kod ve kullanıcıya gösterilecek Türkçe ileti. */
export class SaveError extends Error {
  readonly code: string;
  /** Kısa neden ("Disk dolu"). */
  readonly reason: string;
  constructor(code: string, message?: string, reason?: string) {
    const r = reason || saveErrorReason(code);
    super(message || `${r}.`);
    this.name = "SaveError";
    this.code = code;
    this.reason = r;
  }
}

/** Başarısız yazma sonucunu SaveError'a çevirir. */
export function saveErrorFrom(outcome: { code?: string; reason?: string; message?: string } | null | undefined): SaveError {
  return new SaveError(outcome?.code || "EUNKNOWN", outcome?.message, outcome?.reason);
}

/** Herhangi bir hatanın kısa nedeni (SaveError değilse genel ileti). */
export function describeSaveError(err: unknown): string {
  if (err instanceof SaveError) return err.reason;
  const name = (err as { name?: string; code?: string } | null)?.code ?? (err as { name?: string } | null)?.name;
  return saveErrorReason(name);
}
