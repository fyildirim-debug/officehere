// Dosya yazma hata kodlarının kullanıcıya gösterilen Türkçe karşılıkları.
// Sayfadaki karşılığı: packages/core/src/saveErrors.ts (aynı kodlar, aynı metinler).

const REASONS: Record<string, string> = {
  ENOSPC: "Disk dolu",
  EDQUOT: "Disk kotası doldu",
  EACCES: "Yazma izni yok",
  EPERM: "Yazma izni yok",
  EBUSY: "Dosya başka bir programda açık",
  ETXTBSY: "Dosya başka bir programda açık",
  EROFS: "Disk salt okunur",
  ENOENT: "Klasör bulunamadı",
  ENOTDIR: "Klasör bulunamadı",
  EISDIR: "Bu adda bir klasör var",
  ENAMETOOLONG: "Dosya adı ya da yolu çok uzun",
  EIO: "Disk okuma/yazma hatası",
  EMFILE: "Çok fazla dosya açık",
  ENFILE: "Çok fazla dosya açık",
  EFBIG: "Dosya bu disk için çok büyük",
  EBACKUP: "Özgün dosyanın yedeği alınamadı",
};

const HINTS: Record<string, string> = {
  ENOSPC: "Diskte yer açıp yeniden deneyin ya da başka bir konuma kaydedin.",
  EDQUOT: "Diskte yer açıp yeniden deneyin ya da başka bir konuma kaydedin.",
  EACCES: "Dosya ya da klasör salt okunur olabilir. Başka bir konuma kaydedin.",
  EPERM: "Dosya ya da klasör salt okunur olabilir. Başka bir konuma kaydedin.",
  EBUSY: "Dosyayı kullanan programı kapatıp yeniden deneyin ya da farklı bir adla kaydedin.",
  ETXTBSY: "Dosyayı kullanan programı kapatıp yeniden deneyin ya da farklı bir adla kaydedin.",
  EROFS: "Başka bir konuma kaydedin.",
  ENOENT: "Klasör taşınmış ya da silinmiş olabilir. Farklı Kaydet ile yeni bir konum seçin.",
  ENOTDIR: "Klasör taşınmış ya da silinmiş olabilir. Farklı Kaydet ile yeni bir konum seçin.",
  EBACKUP: "Dosyanın özgün hâli korunamadığı için üzerine yazılmadı.",
};

/** Kısa neden ("Disk dolu"). */
export function saveErrorReason(code: string | undefined): string {
  return (code && REASONS[code]) || "Dosya yazılamadı";
}

/** Neden ve öneri ("Disk dolu. Diskte yer açıp…"). */
export function saveErrorMessage(code: string | undefined): string {
  const hint = code ? HINTS[code] : undefined;
  return hint ? `${saveErrorReason(code)}. ${hint}` : `${saveErrorReason(code)}.`;
}
