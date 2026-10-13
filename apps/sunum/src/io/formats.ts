/**
 * Sunu dosya biçimleri: uzantı + içerik imzasıyla biçim tanıma, okuyucu ve
 * yazıcı seçimi, aç/kaydet süzgeçleri. Yerel biçim .pptx'tir; .ppt ve
 * makrolu/şablon çeşitleri yalnız açılır.
 */
import JSZip from "jszip";
import { extensionOf, isCfb, type FileFilter } from "@officehere/core";
import type { Presentation } from "../model/types";
import { readOdp } from "./odp/reader";
import { writeOdp } from "./odp/writer";
import { readPpt } from "./ppt/reader";
import { readPptx } from "./pptxReader";
import { writePptx } from "./pptxWriter";

export type SunuFormat = "pptx" | "pptm" | "potx" | "potm" | "ppsx" | "ppsm" | "ppt" | "odp" | "otp";
export type SaveFormat = "pptx" | "ppsx" | "potx" | "odp";

export interface FormatInfo {
  label: string;
  /** Şablon: açılınca adsız yeni sunu olur. */
  template?: boolean;
  /** Slayt gösterisi dosyası. */
  show?: boolean;
  /** Makro içerir (makrolar yok sayılır). */
  macro?: boolean;
  /** Eski ikili biçim (uyumluluk modu). */
  legacy?: boolean;
  /** Aynı biçimde yerinde kaydedilebilir mi. */
  writable?: SaveFormat;
}

export const FORMATS: Record<SunuFormat, FormatInfo> = {
  pptx: { label: "Sunu", writable: "pptx" },
  pptm: { label: "Makro İçeren Sunu", macro: true },
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
  odp: { name: "OpenDocument Sunu", extensions: ["odp"] },
  ppsx: { name: "Slayt Gösterisi", extensions: ["ppsx"] },
  potx: { name: "Şablon", extensions: ["potx"] },
};

export const SAVE_ORDER: SaveFormat[] = ["pptx", "odp", "ppsx", "potx"];

/** Farklı Kaydet süzgeçleri: istenen biçim başta. */
export function saveFilters(first: SaveFormat): FileFilter[] {
  return [SAVE_FILTER[first], ...SAVE_ORDER.filter((f) => f !== first).map((f) => SAVE_FILTER[f])];
}

export function isSaveFormat(ext: string | undefined): ext is SaveFormat {
  return !!ext && (SAVE_ORDER as string[]).includes(ext);
}

export function isSunuFormat(ext: string | undefined): ext is SunuFormat {
  return !!ext && ext in FORMATS;
}

export class FormatError extends Error {}

const OOXML_MAIN: [RegExp, SunuFormat][] = [
  [/presentationml\.slideshow\.main/i, "ppsx"],
  [/presentationml\.template\.main/i, "potx"],
  [/powerpoint\.presentation\.macroenabled/i, "pptm"],
  [/powerpoint\.slideshow\.macroenabled/i, "ppsm"],
  [/powerpoint\.template\.macroenabled/i, "potm"],
  [/presentationml\.presentation\.main/i, "pptx"],
];

/** Uzantı ve içerik imzasıyla biçimi belirler (içerik önceliklidir). */
export async function sniffFormat(data: ArrayBuffer | Uint8Array, name: string): Promise<SunuFormat> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const ext = extensionOf(name);
  if (isCfb(bytes)) return "ppt";
  if (bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04) {
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(bytes);
    } catch {
      throw new FormatError("Dosya bozuk; ZIP paketi okunamadı.");
    }
    const mime = (await zip.file("mimetype")?.async("string"))?.trim();
    if (mime) {
      if (mime.startsWith("application/vnd.oasis.opendocument.presentation-template")) return "otp";
      if (mime.startsWith("application/vnd.oasis.opendocument.presentation")) return "odp";
      if (mime.startsWith("application/vnd.oasis.opendocument.text")) throw new FormatError("Bu dosya bir sunu değil; bir OpenDocument metin belgesi.");
      if (mime.startsWith("application/vnd.oasis.opendocument.spreadsheet")) throw new FormatError("Bu dosya bir sunu değil; bir OpenDocument hesap tablosu.");
    }
    const ct = await zip.file("[Content_Types].xml")?.async("string");
    if (ct) {
      const fam = ext === "pptm" || ext === "ppsx" || ext === "ppsm" || ext === "potx" || ext === "potm" ? ext : "pptx";
      // Genel sunu içerik türünde (bazı yazıcıların .ppsx'i) uzantı belirleyicidir.
      for (const [re, f] of OOXML_MAIN) if (re.test(ct)) return f === "pptx" ? fam : f;
      if (/wordprocessingml/i.test(ct)) throw new FormatError("Bu dosya bir sunu değil; bir kelime işlemci belgesi.");
      if (/spreadsheetml/i.test(ct)) throw new FormatError("Bu dosya bir sunu değil; bir hesap tablosu.");
      return fam;
    }
    if (zip.file("content.xml")) return ext === "otp" ? "otp" : "odp";
    throw new FormatError("Dosya biçimi tanınmadı (bilinen bir sunu paketi değil).");
  }
  throw new FormatError("Dosya biçimi tanınmadı; desteklenen bir sunu dosyası değil.");
}

export interface ReadResult {
  pres: Presentation;
  format: SunuFormat;
}

/** Biçimi tanıyıp uygun okuyucuyla açar. */
export async function readPresentationFile(data: ArrayBuffer | Uint8Array, name: string): Promise<ReadResult> {
  const format = await sniffFormat(data, name);
  let pres: Presentation;
  if (format === "ppt") pres = await readPpt(data);
  else if (format === "odp" || format === "otp") pres = await readOdp(data);
  else pres = await readPptx(data);
  return { pres, format };
}

/** Sunuyu seçilen biçimde yazar. */
export function writePresentationFormat(pres: Presentation, format: SaveFormat): Promise<Uint8Array> {
  switch (format) {
    case "odp":
      return writeOdp(pres);
    case "ppsx":
      return writePptx(pres, { kind: "slideshow" });
    case "potx":
      return writePptx(pres, { kind: "template" });
    default:
      return writePptx(pres);
  }
}
