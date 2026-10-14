// Sayfa ile Worker'ın ortak hata sınıfı (uygulamalardaki FormatError'ın benzeri).
export class FormatError extends Error {
  detail?: string;
  constructor(message: string, detail?: string) {
    super(message);
    this.name = "FormatError";
    if (detail) this.detail = detail;
  }
}
