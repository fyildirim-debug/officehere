/**
 * Sunu dosya biçimlerinin ağır kısmı: uzantı + içerik imzasıyla biçim tanıma,
 * okuyucu ve yazıcı seçimi. JSZip ve tüm okuyucu/yazıcılarla birlikte ayrı bir
 * parçadadır; uygulama buna formats.ts üzerinden (tembel) erişir.
 */
import { extensionOf, isCfb } from "@officehere/core";
import type { Presentation } from "../model/types";
import { FormatError, toFormatError, type SaveFormat, type SunuFormat } from "./formatInfo";
import { LoadWarnings } from "./loadWarnings";
import { readOdp } from "./odp/reader";
import { writeOdp } from "./odp/writer";
import { hasOdfShapes, ODF_OBJECTS_FALLBACK } from "./odp/objects";
import { readPpt } from "./ppt/reader";
import { readPptx } from "./pptxReader";
import { writePptx } from "./pptxWriter";
import { hasRawShapes, RAW_WARN } from "./rawPassthrough";
import { openZip, PackageError, PartError, type SafeZip } from "./zipSafe";

/** JSZip, bu tembel parçanın içinden dışa verilir (ör. çok slaytlı PNG dışa aktarma, anahat içe aktarma). */
export { default as JSZip } from "jszip";

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
    let zip: SafeZip;
    try {
      zip = await openZip(bytes, "Dosya bozuk; ZIP paketi okunamadı.");
    } catch (e) {
      throw new FormatError((e as Error).message, (e as PackageError).detail);
    }
    // Tanıma parçaları bozuksa (ör. bayt bozulması) yok sayılır; biçim parça adlarından çıkarılır.
    const small = async (p: string) => {
      try {
        return await zip.text(p, 4 * 1024 * 1024);
      } catch (e) {
        if (e instanceof PartError) return undefined;
        throw e;
      }
    };
    const mime = (await small("mimetype"))?.trim();
    if (mime) {
      if (mime.startsWith("application/vnd.oasis.opendocument.presentation-template")) return "otp";
      if (mime.startsWith("application/vnd.oasis.opendocument.presentation")) return "odp";
      if (mime.startsWith("application/vnd.oasis.opendocument.text")) throw new FormatError("Bu dosya bir sunu değil; bir OpenDocument metin belgesi.");
      if (mime.startsWith("application/vnd.oasis.opendocument.spreadsheet")) throw new FormatError("Bu dosya bir sunu değil; bir OpenDocument hesap tablosu.");
    }
    const ct = (await small("[Content_Types].xml")) ?? (zip.has("ppt/presentation.xml") ? "presentationml.presentation.main" : undefined);
    if (ct) {
      const fam = ext === "pptm" || ext === "ppsx" || ext === "ppsm" || ext === "potx" || ext === "potm" ? ext : "pptx";
      // Genel sunu içerik türünde (bazı yazıcıların .ppsx'i) uzantı belirleyicidir.
      for (const [re, f] of OOXML_MAIN) if (re.test(ct)) return f === "pptx" ? fam : f;
      if (/wordprocessingml/i.test(ct)) throw new FormatError("Bu dosya bir sunu değil; bir kelime işlemci belgesi.");
      if (/spreadsheetml/i.test(ct)) throw new FormatError("Bu dosya bir sunu değil; bir hesap tablosu.");
      return fam;
    }
    if (zip.has("content.xml")) return ext === "otp" ? "otp" : "odp";
    throw new FormatError("Dosya biçimi tanınmadı (bilinen bir sunu paketi değil).");
  }
  throw new FormatError("Dosya biçimi tanınmadı; desteklenen bir sunu dosyası değil.");
}

export interface ReadResult {
  pres: Presentation;
  format: SunuFormat;
  /**
   * Açılış uyarıları (Türkçe, kullanıcıya gösterilir): okunamayan, eksik ya da
   * yerine boş slayt konan bölümler. Boş değilse dosya tam açılmamıştır;
   * otomatik kaydetme bu dosyanın üzerine yazmamalıdır.
   */
  warnings: string[];
  /** Uyarıların teknik ayrıntıları (parça yolları, özgün hata iletileri). */
  details: string[];
  /**
   * Hasar olmayan bilgi iletileri (Türkçe): ör. dijital imza kaydedince kaldırılır. Otomatik
   * kaydetme bunlar onaylanmadan dosyanın üzerine yazmamalıdır.
   */
  notices: string[];
}

export interface ReadOptions {
  /** Özgün paket geçiş katmanı için kaydedilsin mi (varsayılan evet; tek kullanımlık doğrulama okumasında hayır). */
  kaynak?: boolean;
}

export interface WriteOptions {
  /** Kaydetme uyarıları (Türkçe): düz gruba çevrilen SmartArt, kaldırılan imza/makro, ODP yedek görünümü… */
  warnings?: string[];
}

/** Biçimi tanıyıp uygun okuyucuyla açar. Hatalar Türkçe FormatError olarak atılır. */
export async function readPresentationFile(data: ArrayBuffer | Uint8Array, name: string, opts: ReadOptions = {}): Promise<ReadResult> {
  const warnings = new LoadWarnings();
  const notices: string[] = [];
  try {
    const format = await sniffFormat(data, name);
    let pres: Presentation;
    if (format === "ppt") pres = await readPpt(data, { warnings });
    else if (format === "odp" || format === "otp") pres = await readOdp(data, { warnings, kaynak: opts.kaynak });
    else pres = await readPptx(data, { warnings, notices, kaynak: opts.kaynak });
    return { pres, format, warnings: warnings.list(), details: warnings.details, notices };
  } catch (e) {
    throw toFormatError(e);
  }
}

/** Sunuyu seçilen biçimde yazar. */
export function writePresentationFormat(pres: Presentation, format: SaveFormat, opts: WriteOptions = {}): Promise<Uint8Array> {
  const { warnings } = opts;
  if (format === "odp") {
    // .pptx'in ham öğeleri ODF'de yedek görünümleriyle (şekil grubu/resim) yazılır; .odp'den gelen gömülü
    // nesneler özgün çerçeveleriyle (ODF geçiş katmanı, odp/writer.ts).
    if (warnings && hasRawShapes([...pres.slides, ...pres.layouts, pres.master])) warnings.push(RAW_WARN.odp);
    return writeOdp(pres, { warnings });
  }
  // .odp'ye özgü gömülü nesneler .pptx'te yedek görünümleriyle (resim) yazılır.
  if (warnings && hasOdfShapes([...pres.slides, ...pres.layouts, pres.master])) warnings.push(ODF_OBJECTS_FALLBACK);
  switch (format) {
    case "ppsx":
      return writePptx(pres, { kind: "slideshow", warnings });
    case "potx":
      return writePptx(pres, { kind: "template", warnings });
    case "pptm":
      return writePptx(pres, { macro: true, warnings });
    default:
      return writePptx(pres, { warnings });
  }
}
