/**
 * Dosya biçimleri: içerik imzası + uzantıyla biçim tanıma, açma ve yazıcı seçimi.
 * Yerel biçim .xlsx'tir; .xls ve şablonlar yalnız açılır, .ods/.csv/.tsv yazılabilir.
 */
import JSZip from "jszip";
import { baseName, extensionOf, isCfb, type FileFilter } from "@officehere/core";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import { FormatError } from "./bookUtil";
import { readCsv, writeCsv } from "./csv";
import { readHtml } from "./html";
import { ODS_MIME, readOds } from "./ods/reader";
import { writeOds } from "./ods/writer";
import { readXls } from "./xls/reader";
import { readXlsx, writeXlsx } from "./xlsx";

export type BookFormat = "xlsx" | "xlsm" | "xltx" | "xltm" | "xls" | "ods" | "ots" | "csv" | "tsv" | "txt" | "html";
/** Kaydedilebilen biçimler. */
export type SaveFormat = "xlsx" | "ods" | "csv" | "tsv";

export const FILTERS: Record<SaveFormat, FileFilter> = {
  xlsx: { name: "Çalışma Kitabı (.xlsx)", extensions: ["xlsx"] },
  ods: { name: "OpenDocument Hesap Tablosu (.ods)", extensions: ["ods"] },
  csv: { name: "CSV UTF-8 (Noktalı virgülle ayrılmış, yalnız etkin sayfa) (.csv)", extensions: ["csv"] },
  tsv: { name: "Metin (Sekmeyle ayrılmış, yalnız etkin sayfa) (.tsv)", extensions: ["tsv"] },
};

/** Aç iletişim kutusu filtreleri: tümü + biçim grupları. */
export const OPEN_FILTERS: FileFilter[] = [
  { name: "Tüm Çalışma Kitapları", extensions: ["xlsx", "xlsm", "xltx", "xltm", "xls", "ods", "ots", "csv", "tsv", "txt", "html", "htm"] },
  { name: "Çalışma Kitabı (.xlsx, .xlsm)", extensions: ["xlsx", "xlsm"] },
  { name: "Şablon (.xltx, .xltm, .ots)", extensions: ["xltx", "xltm", "ots"] },
  { name: "97-2003 Çalışma Kitabı (.xls)", extensions: ["xls"] },
  { name: "OpenDocument Hesap Tablosu (.ods)", extensions: ["ods", "ots"] },
  { name: "CSV (Virgülle/noktalı virgülle ayrılmış) (.csv)", extensions: ["csv"] },
  { name: "Metin (Sekmeyle ayrılmış) (.tsv, .txt)", extensions: ["tsv", "txt"] },
  { name: "Web Sayfası (.html)", extensions: ["html", "htm"] },
];

const OOXML_MAIN: Record<string, BookFormat> = {
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "xlsx",
  "application/vnd.ms-excel.sheet.macroEnabled.main+xml": "xlsm",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml": "xltx",
  "application/vnd.ms-excel.template.macroEnabled.main+xml": "xltm",
};

function startsWith(b: Uint8Array, sig: number[]): boolean {
  return sig.every((v, i) => b[i] === v);
}

/** İçerik imzasına (yoksa uzantıya) göre biçimi bulur. */
export async function sniffFormat(name: string, data: ArrayBuffer): Promise<BookFormat> {
  const ext = extensionOf(name);
  const head = new Uint8Array(data, 0, Math.min(data.byteLength, 1024));
  if (isCfb(new Uint8Array(data))) return "xls";
  if (startsWith(head, [0x50, 0x4b, 0x03, 0x04])) {
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(data);
    } catch {
      throw new FormatError("Dosya bozuk: ZIP paketi okunamadı.");
    }
    const mime = (await zip.file("mimetype")?.async("string"))?.trim();
    if (mime) {
      if (mime === ODS_MIME) return "ods";
      if (mime === `${ODS_MIME}-template`) return "ots";
      if (mime.includes("opendocument.text")) throw new FormatError("Bu bir OpenDocument metin belgesi; OfficeHere Belge ile açın.");
      if (mime.includes("opendocument.presentation")) throw new FormatError("Bu bir OpenDocument sunusu; OfficeHere Sunum ile açın.");
    }
    const ct = await zip.file("[Content_Types].xml")?.async("string");
    if (ct) {
      for (const [type, fmt] of Object.entries(OOXML_MAIN)) if (ct.includes(`"${type}"`)) return fmt;
      if (ct.includes("wordprocessingml")) throw new FormatError("Bu bir belge dosyası; OfficeHere Belge ile açın.");
      if (ct.includes("presentationml")) throw new FormatError("Bu bir sunu dosyası; OfficeHere Sunum ile açın.");
      return "xlsx";
    }
    if (zip.file("content.xml")) return "ods";
    return ext === "ods" || ext === "ots" ? (ext as BookFormat) : "xlsx";
  }
  // Metin biçimleri
  const text = new TextDecoder("utf-8").decode(head).replace(/^﻿/, "").trimStart().toLowerCase();
  if (ext === "html" || ext === "htm" || /^<(!doctype html|html|table|meta|head|body)\b/.test(text)) return "html";
  if (ext === "tsv") return "tsv";
  if (ext === "txt") return "txt";
  if (ext === "xlsx" || ext === "xls" || ext === "ods") throw new FormatError("Dosya içeriği uzantısıyla uyuşmuyor ya da dosya bozuk.");
  return "csv";
}

export interface OpenedBook {
  wb: WorkbookModel;
  format: BookFormat;
  /** Şablondan açıldı: adsız yeni çalışma kitabı gibi davranır. */
  template: boolean;
  /** İleti çubuğunda gösterilecek bilgi. */
  notice?: string;
}

/** Dosyayı biçimine uygun okuyucuyla açar. */
export async function openBook(name: string, data: ArrayBuffer): Promise<OpenedBook> {
  const format = await sniffFormat(name, data);
  const sheetName = baseName(name);
  switch (format) {
    case "xls":
      return { wb: readXls(data), format, template: false, notice: "Uyumluluk modu: bu dosya eski 97-2003 biçiminde; bazı öğeler gösterilemeyebilir. Kaydedince .xlsx olarak kaydedilir." };
    case "ods":
    case "ots":
      return { wb: await readOds(data), format, template: format === "ots" };
    case "csv":
      return { wb: readCsv(data, sheetName), format, template: false };
    case "tsv":
      return { wb: readCsv(data, sheetName, { delimiter: "\t" }), format, template: false };
    case "txt":
      return { wb: readCsv(data, sheetName), format, template: false };
    case "html":
      return { wb: readHtml(data, sheetName), format, template: false, notice: "Web sayfasındaki tablolar içe aktarıldı. Kaydedince .xlsx olarak kaydedilir." };
    default: {
      const wb = await readXlsx(data);
      const macro = format === "xlsm" || format === "xltm";
      return { wb, format, template: format === "xltx" || format === "xltm", notice: macro ? "Bu çalışma kitabındaki makrolar çalıştırılmaz ve kaydedilirken korunmaz. Kaydedince .xlsx olarak kaydedilir." : undefined };
    }
  }
}

/** Biçim aynı biçimde yazılabiliyor mu (yoksa Kaydet → Farklı Kaydet .xlsx). */
export function writableAs(format: BookFormat): SaveFormat | null {
  if (format === "xlsx" || format === "ods" || format === "csv" || format === "tsv") return format;
  if (format === "txt") return "tsv";
  return null;
}

/** Uzantıdan kaydetme biçimi. */
export function saveFormatOf(name: string): SaveFormat | null {
  const e = extensionOf(name);
  return e === "xlsx" || e === "ods" || e === "csv" || e === "tsv" ? e : e === "txt" ? "tsv" : null;
}

/** Çalışma kitabını istenen biçimde baytlara/metne çevirir. */
export async function writeBook(wb: WorkbookModel, sheet: SheetModel, format: SaveFormat, delimiter?: string): Promise<Uint8Array<ArrayBuffer> | string> {
  if (format === "csv") return writeCsv(wb, sheet, { delimiter: delimiter ?? ";" });
  if (format === "tsv") return writeCsv(wb, sheet, { delimiter: delimiter ?? "\t" });
  if (format === "ods") return writeOds(wb);
  return writeXlsx(wb);
}
