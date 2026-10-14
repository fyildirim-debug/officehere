/**
 * Açılış hatalarının sınıflandırılması (JSZip ve okuyuculara bağımlı değil):
 * ilk yüklemede ağır G/Ç parçası olmadan da kullanılabilir.
 */

/** Kullanıcıya gösterilecek (Türkçe) ileti + günlük için teknik ayrıntı. */
export class PackageError extends Error {
  detail?: string;
  constructor(message: string, detail?: string) {
    super(message);
    this.name = "PackageError";
    this.detail = detail;
  }
}

/** Teknik hata iletisi (JSZip/pako/tarayıcı) → Türkçe sınıf iletisi. */
export function describeZipError(e: unknown): string {
  const m = String((e as Error)?.message ?? e);
  if (/encrypt/i.test(m)) return "Dosya parola ile korunuyor (şifreli ZIP); açılamıyor.";
  if (/compression|unsupported/i.test(m)) return "Dosya desteklenmeyen bir sıkıştırma yöntemi kullanıyor.";
  if (/string length|allocation|memory|array buffer/i.test(m)) return "Dosya bu ortamda açılamayacak kadar büyük.";
  return "Dosya bozuk; sunu paketi okunamadı.";
}
