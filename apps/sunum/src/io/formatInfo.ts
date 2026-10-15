/**
 * Sunu dosya biçimleri bilgisi (hafif): biçim tablosu, aç/kaydet süzgeçleri,
 * biçim denetimleri ve hata sınıflandırması. Okuyucu/yazıcılar (JSZip ile
 * birlikte) ayrı bir parçadadır ve ilk kullanımda yüklenir (formats.ts).
 */
import type { FileFilter } from "@officehere/core";
import { describeZipError, PackageError } from "./errors";

export type SunuFormat = "pptx" | "pptm" | "potx" | "potm" | "ppsx" | "ppsm" | "ppt" | "odp" | "otp";
export type SaveFormat = "pptx" | "pptm" | "ppsx" | "potx" | "odp";

export interface FormatInfo {
  label: string;
  /** Şablon: açılınca adsız yeni sunu olur. */
  template?: boolean;
  /** Slayt gösterisi dosyası. */
  show?: boolean;
  /** Makro içerir (makrolar çalıştırılmaz; .pptm kaydederken korunur). */
  macro?: boolean;
  /** Eski ikili biçim (uyumluluk modu). */
  legacy?: boolean;
  /** Aynı biçimde yerinde kaydedilebilir mi. */
  writable?: SaveFormat;
}

export const FORMATS: Record<SunuFormat, FormatInfo> = {
  pptx: { label: "Sunu", writable: "pptx" },
  pptm: { label: "Makro İçeren Sunu", macro: true, writable: "pptm" },
  potx: { label: "Şablon", template: true, writable: "potx" },
  potm: { label: "Makro İçeren Şablon", template: true, macro: true },
  ppsx: { label: "Slayt Gösterisi", show: true, writable: "ppsx" },
  ppsm: { label: "Makro İçeren Slayt Gösterisi", show: true, macro: true },
  ppt: { label: "97-2003 Sunu", legacy: true },
  odp: { label: "OpenDocument Sunu", writable: "odp" },
  otp: { label: "OpenDocument Sunu Şablonu", template: true },
};

export const OPEN_EXTENSIONS = ["pptx", "pptm", "ppt", "odp", "ppsx", "ppsm", "potx", "potm", "otp"];

/** Aç iletişim kutusu süzgeçleri. */
export const OPEN_FILTERS: FileFilter[] = [
  { name: "Tüm Sunular", extensions: OPEN_EXTENSIONS },
  { name: "Sunu", extensions: ["pptx", "pptm"] },
  { name: "97-2003 Sunu", extensions: ["ppt"] },
  { name: "OpenDocument Sunu", extensions: ["odp"] },
  { name: "Slayt Gösterisi", extensions: ["ppsx", "ppsm"] },
  { name: "Şablon", extensions: ["potx", "potm", "otp"] },
];

const SAVE_FILTER: Record<SaveFormat, FileFilter> = {
  pptx: { name: "Sunu", extensions: ["pptx"] },
  pptm: { name: "Makro İçeren Sunu", extensions: ["pptm"] },
  odp: { name: "OpenDocument Sunu", extensions: ["odp"] },
  ppsx: { name: "Slayt Gösterisi", extensions: ["ppsx"] },
  potx: { name: "Şablon", extensions: ["potx"] },
};

export const SAVE_ORDER: SaveFormat[] = ["pptx", "odp", "ppsx", "potx"];

/** Farklı Kaydet listesinde olmayan ama yerinde yazılabilen biçimler (makrolu sunu yalnız kendi dosyasına). */
const IN_PLACE_ONLY: SaveFormat[] = ["pptm"];

/** Farklı Kaydet süzgeçleri: istenen biçim başta. */
export function saveFilters(first: SaveFormat): FileFilter[] {
  return [SAVE_FILTER[first], ...SAVE_ORDER.filter((f) => f !== first).map((f) => SAVE_FILTER[f])];
}

export function isSaveFormat(ext: string | undefined): ext is SaveFormat {
  return !!ext && ((SAVE_ORDER as string[]).includes(ext) || (IN_PLACE_ONLY as string[]).includes(ext));
}

export function isSunuFormat(ext: string | undefined): ext is SunuFormat {
  return !!ext && ext in FORMATS;
}

/** Biçim/açılış hatası: Türkçe ileti; `detail` günlük ve "Ayrıntılar" için teknik ayrıntı. */
export class FormatError extends Error {
  detail?: string;
  constructor(message: string, detail?: string) {
    super(message);
    this.name = "FormatError";
    this.detail = detail;
  }
}

/** Okuyuculardan gelen her hatayı Türkçe, sınıflandırılmış bir FormatError'a çevirir. */
export function toFormatError(e: unknown): FormatError {
  if (e instanceof FormatError) return e;
  const raw = String((e as Error)?.message ?? e);
  const detail = (e as { detail?: string })?.detail ?? `${(e as Error)?.name ?? "Error"}: ${raw}`;
  // Okuyucuların kendi (Türkçe) iletileri olduğu gibi kullanılır.
  if (e instanceof PackageError || /[çğıöşüÇĞİÖŞÜ]/.test(raw)) return new FormatError(raw, (e as PackageError).detail);
  if (e instanceof RangeError && /call stack/i.test(raw)) return new FormatError("Dosya olağandışı derin iç içe yapı içeriyor; açılamadı. Dosya bozuk ya da zararlı olabilir.", detail);
  if (/string length|allocation|out of memory|array buffer/i.test(raw)) return new FormatError("Dosya bu ortamda açılamayacak kadar büyük.", detail);
  if (/zip|central directory|signature|compress|inflate|deflate|crc|size mismatch|corrupt|end of data|invalid (block|distance|code|stored)/i.test(raw)) return new FormatError(describeZipError(e), detail);
  if (/dynamically imported module|importing a module script failed|error loading dynamically imported|ChunkLoadError|unable to preload/i.test(raw)) return new FormatError(LOAD_FAILED, detail);
  return new FormatError("Dosya bozuk ya da beklenmeyen bir yapıda; okunamadı.", detail);
}

/** Uygulama parçası (ör. dosya okuyucu/yazıcı) yüklenemedi. */
export const LOAD_FAILED = "Uygulamanın dosya bileşeni yüklenemedi. Bağlantınızı denetleyip yeniden deneyin ya da sayfayı yenileyin.";

/** Dijital imzalı .pptx açılış bildirimi (ReadResult.notices). */
export const SIGNED_NOTICE = "Bu sunu dijital olarak imzalı; kaydederseniz imza kaldırılır.";
