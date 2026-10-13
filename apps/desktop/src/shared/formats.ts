// Masaüstünün tanıdığı dosya biçimleri: uzantı → açacak uygulama.
// Ana süreç (komut satırı, çift tık, oturum) ve başlangıç ekranı (Aç iletişim kutusu) birlikte kullanır.

export type FormatApp = "belge" | "tablo" | "sunum";

export interface FormatGroup {
  app: FormatApp;
  name: string;
  extensions: string[];
}

/** Uygulama başına biçim grupları (iletişim kutusu filtre sırası). */
export const FORMAT_GROUPS: FormatGroup[] = [
  { app: "belge", name: "Belgeler", extensions: ["docx", "docm", "dotx", "dotm", "doc", "odt", "ott", "rtf", "md", "markdown", "txt", "html", "htm"] },
  { app: "tablo", name: "Çalışma kitapları", extensions: ["xlsx", "xlsm", "xltx", "xltm", "xls", "ods", "ots", "csv", "tsv"] },
  { app: "sunum", name: "Sunular", extensions: ["pptx", "pptm", "potx", "potm", "ppsx", "ppsm", "ppt", "odp", "otp"] },
];

const EXT_APP: Record<string, FormatApp> = Object.fromEntries(FORMAT_GROUPS.flatMap((g) => g.extensions.map((e) => [e, g.app])));

/** Dosya adı ya da yolundan uygulamayı bulur; tanınmayan uzantıda null. */
export function appForFileName(name: string): FormatApp | null {
  const base = name.split(/[\\/]/).pop() ?? "";
  const i = base.lastIndexOf(".");
  return i < 0 ? null : (EXT_APP[base.slice(i + 1).toLowerCase()] ?? null);
}

/** Tüm tanınan uzantılar. */
export const ALL_EXTENSIONS = Object.keys(EXT_APP);
