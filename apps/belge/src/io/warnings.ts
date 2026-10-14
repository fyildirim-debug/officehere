// Açılış uyarıları ve hata sınıflandırması: okuyucunun okuyamadığı (kaydedince kaybolacak) içerik
// türleri ileti çubuğunda Türkçe gösterilir; ham kitaplık hataları anlaşılır iletiye çevrilir.

/** Okunamayan ya da tam okunamayan (ya da kaydedince korunamayan) içerik türleri. */
export type LossKind =
  | "signature"
  | "trackChanges"
  | "trackFormat"
  | "comments"
  | "commentThreads"
  | "fields"
  | "objects"
  | "shapes"
  | "macros"
  | "images"
  | "corrupt"
  | "unclosed"
  | "deep"
  | "simplified"
  | "flattened";

const MESSAGES: Record<LossKind, string> = {
  signature: "Bu belge dijital olarak imzalanmış. Kaydederseniz dijital imza kaldırılır.",
  trackChanges: "Bu belgedeki değişiklik izleme kayıtları bu biçimde okunamadı; kaydederseniz kaybolur.",
  trackFormat: "Bu belgedeki bazı değişiklik izleme kayıtları (paragraf, tablo ya da bölüm biçimi değişiklikleri) okunamadı; kaydederseniz kaybolur.",
  comments: "Bu belgedeki açıklamalar okunamadı; kaydederseniz kaybolur.",
  commentThreads: "Açıklama yanıtları ayrı açıklamalar olarak açıldı; kaydederseniz yanıt bağları ve \"çözüldü\" işaretleri kaybolur.",
  fields: "Bazı alanlar yalnız görünen değerleriyle açıldı; kaydederseniz alan olarak değil düz metin olarak kalır.",
  objects: "Belgedeki gömülü nesneler okunamadı ya da yalnız görüntüleriyle açıldı; kaydederseniz nesnelerin kendisi kaybolur.",
  shapes: "Belgedeki bazı metin kutuları ya da çizimler okunamadı; kaydederseniz kaybolur.",
  macros: "Bu belgedeki makrolar çalıştırılmaz. Belge .docm olarak kaydedilince makrolar korunur; .docx ya da başka biçimde kaydedilirse kaldırılır.",
  images: "Bazı resimler bulunamadı ya da okunamadı; kaydederseniz bu resimler kaybolur.",
  corrupt: "Dosyanın bazı bölümleri bozuk olduğu için atlandı; kaydederseniz bu bölümler kaybolur.",
  unclosed: "Dosya bozuk görünüyor (yapı eksik ya da kesik); içeriğin bir kısmı okunamamış olabilir. Kaydederseniz okunamayan kısım kaybolur.",
  deep: "Belgede aşırı iç içe yapı vardı; derin düzeyler düzleştirilerek açıldı ve kaydederseniz bu yapı kaybolur.",
  simplified: "Belge basitleştirilmiş modda açıldı; biçimlerin bir kısmı okunamadı ve kaydederseniz kaybolur.",
  flattened: "Belgenin yapısı tam okunamadı; metin biçimsiz olarak açıldı. Kaydederseniz biçimler kaybolur.",
};

/** Okuyucuların doldurduğu uyarı kümesi (her tür bir kez; ayrıntı günlük içindir). */
export class LoadWarnings {
  private kinds = new Set<LossKind>();
  private extra: string[] = [];
  readonly details: string[] = [];

  add(kind: LossKind, detail?: string): void {
    this.kinds.add(kind);
    if (detail && this.details.length < 50 && !this.details.includes(detail)) this.details.push(detail);
  }

  /** Hazır türlerin dışında kalan ileti (ör. çözülemeyen resim yolu). */
  addText(text: string): void {
    if (!this.extra.includes(text)) this.extra.push(text);
  }

  has(kind: LossKind): boolean {
    return this.kinds.has(kind);
  }

  get size(): number {
    return this.kinds.size + this.extra.length;
  }

  /** Kullanıcıya gösterilecek iletiler (sıralı, yinelenmesiz). */
  messages(): string[] {
    const order = Object.keys(MESSAGES) as LossKind[];
    return [...order.filter((k) => this.kinds.has(k)).map((k) => MESSAGES[k]), ...this.extra];
  }
}

export function lossMessage(kind: LossKind): string {
  return MESSAGES[kind];
}

/* ------------------------------------------------------------------ */
/* Kaydetme (geçiş katmanı) bildirimleri                                */
/* ------------------------------------------------------------------ */

/** Kurtarılan taslağın özgün paketi bulunamadı (kaynak deposunda yok): geçiş katmanı olmadan kaydedilir. */
export const SOURCE_MISSING =
  "Belgenin özgün dosyası kurtarma deposunda bulunamadı. Belge kurtarılan içerikle açıldı; kaydederseniz OfficeHere'in okuyamadığı parçalar (tema, gömülü yazı tipleri, makrolar, desteklenmeyen öğeler) korunamaz.";

/** Makrolu belge makrosuz biçimde kaydedildi. */
export const MACROS_REMOVED = "Makrolar .docx biçiminde saklanamadığı için bu kayıtta kaldırıldı. Korumak için belgeyi .docm olarak kaydedin.";

/**
 * Paket düzeyi taşıma raporundan kaydederken bildirilecekler. Açılışta uyarılanlar (dijital imza, açıklama yanıt
 * zinciri) her kayıtta yinelenmez.
 */
export function saveNotes(report: { dropped: readonly { reason: string }[] } | null | undefined): string[] {
  return report?.dropped.some((d) => d.reason === "macro") ? [MACROS_REMOVED] : [];
}

/* ------------------------------------------------------------------ */
/* Hata sınıflandırması                                                 */
/* ------------------------------------------------------------------ */

/** Belge açılamadığında fırlatılan anlaşılır hata; teknik ayrıntı `detail` alanındadır. */
export class BelgeOpenError extends Error {
  readonly detail: string;
  constructor(message: string, detail = "") {
    super(message);
    this.name = "FormatError";
    this.detail = detail;
  }
}

/** Kendi okuyucularımızın Türkçe iletisi mi (zaten anlaşılır). */
function isOurMessage(msg: string): boolean {
  return /[çğıöşüÇĞİÖŞÜ]/.test(msg) && !/^\s*(Error|TypeError|RangeError)\b/.test(msg);
}

const LABEL: Record<string, string> = { docx: "belge", odt: "OpenDocument metin belgesi", rtf: "RTF belgesi", doc: ".doc belgesi", md: "Markdown belgesi", html: "web sayfası", txt: "metin dosyası" };

/** Ham hata → anlaşılır Türkçe ileti (ayrıntı ayrı tutulur). */
export function classifyOpenError(err: unknown, format: string): BelgeOpenError {
  if (err instanceof BelgeOpenError) return err;
  const e = err as Error | undefined;
  const raw = String(e?.message ?? err ?? "");
  const detail = `${e?.name && e.name !== "Error" ? `${e.name}: ` : ""}${raw}`.slice(0, 500);
  const what = LABEL[format] ?? "belge";
  if (/^XML ayrıştırılamadı/.test(raw)) return new BelgeOpenError(`Dosya bozuk: ${what} içeriği (XML) okunamadı.`, detail);
  // Okuyucularımızın kendi Türkçe iletisi zaten anlaşılırdır.
  if (isOurMessage(raw)) return new BelgeOpenError(raw, detail);
  if (e?.name === "RangeError" && /call stack/i.test(raw)) return new BelgeOpenError(`Dosya açılamadı: ${what} aşırı iç içe yapı içeriyor ve güvenle okunamıyor.`, detail);
  if (/zip|central directory|end of data|uncompressed data size|unexpected signature|missing \d+ bytes|compression|corrupted|invalid (stored|block|distance|code)|crc32|is this a zip/i.test(raw))
    return new BelgeOpenError("Dosya bozuk: sıkıştırılmış paket (ZIP) okunamadı. Dosya eksik indirilmiş ya da hasar görmüş olabilir.", detail);
  if (/xmldom|parse ?error|not well-formed|unbound namespace|unexpected (close|end)|attribute|entity|root element|only one element|invalid doc source|XML ayrıştırılamadı|Hierarchy request|text data outside of root/i.test(raw))
    return new BelgeOpenError(`Dosya bozuk: ${what} içeriği (XML) okunamadı.`, detail);
  if (/body element|main document part|document\.xml|are you sure this is a/i.test(raw)) return new BelgeOpenError(`Dosya bozuk ya da geçerli bir ${what} değil: ana belge bölümü bulunamadı.`, detail);
  if (/out of memory|allocation failed|array buffer allocation|invalid (array|string) length/i.test(raw)) return new BelgeOpenError("Dosya açılamadı: içerik bu cihazda işlenemeyecek kadar büyük.", detail);
  return new BelgeOpenError(`Dosya açılamadı: ${what} biçimi geçersiz ya da dosya bozuk.`, detail);
}
