// Belge dosya biçimleri: uzantılar, iletişim kutusu filtreleri ve içerik imzasıyla biçim tanıma.
import JSZip from "jszip";
import { extensionOf, isCfb, type FileFilter } from "@officehere/core";

export type FormatId = "docx" | "doc" | "odt" | "fodt" | "rtf" | "md" | "html" | "txt";

export interface FormatDef {
  id: FormatId;
  /** Filtre ve biçim sorusu adı. */
  label: string;
  /** Kısa ad ("OpenDocument Metin"). */
  short: string;
  extensions: string[];
  /** Kaydederken kullanılan uzantı (yazılamayan biçimde null). */
  saveExt: string | null;
}

export const FORMATS: Record<FormatId, FormatDef> = {
  docx: { id: "docx", label: "Belge (.docx)", short: "Belge", extensions: ["docx", "docm", "dotx", "dotm"], saveExt: "docx" },
  doc: { id: "doc", label: "Belge 97-2003 (.doc)", short: ".doc (97-2003)", extensions: ["doc"], saveExt: null },
  odt: { id: "odt", label: "OpenDocument Metin (.odt)", short: "OpenDocument Metin", extensions: ["odt", "ott"], saveExt: "odt" },
  fodt: { id: "fodt", label: "Düz XML OpenDocument Metin (.fodt)", short: "Düz XML OpenDocument", extensions: ["fodt"], saveExt: null },
  rtf: { id: "rtf", label: "Zengin Metin (.rtf)", short: "Zengin Metin", extensions: ["rtf"], saveExt: "rtf" },
  md: { id: "md", label: "Markdown (.md)", short: "Markdown", extensions: ["md", "markdown"], saveExt: "md" },
  html: { id: "html", label: "Web Sayfası (.html)", short: "Web Sayfası", extensions: ["html", "htm"], saveExt: "html" },
  txt: { id: "txt", label: "Düz Metin (.txt)", short: "Düz Metin", extensions: ["txt"], saveExt: "txt" },
};

const OPEN_ORDER: FormatId[] = ["docx", "doc", "odt", "fodt", "rtf", "md", "html", "txt"];
const SAVE_ORDER: FormatId[] = ["docx", "odt", "rtf", "md", "html", "txt"];

export const OPEN_FILTERS: FileFilter[] = [
  { name: "Tüm Belgeler", extensions: OPEN_ORDER.flatMap((f) => FORMATS[f].extensions) },
  ...OPEN_ORDER.map((f) => ({ name: FORMATS[f].label, extensions: FORMATS[f].extensions })),
];

/** <input type=file> accept değeri. */
export const OPEN_ACCEPT = OPEN_ORDER.flatMap((f) => FORMATS[f].extensions.map((e) => `.${e}`)).join(",");

export const DOCX_FILTER: FileFilter[] = [{ name: FORMATS.docx.label, extensions: ["docx"] }];

/** Farklı Kaydet biçim listesi (verilen biçim başta). */
export function saveFilters(first: FormatId = "docx"): FileFilter[] {
  const order = [first, ...SAVE_ORDER.filter((f) => f !== first)].filter((f) => FORMATS[f].saveExt);
  return order.map((f) => ({ name: FORMATS[f].label, extensions: [FORMATS[f].saveExt!] }));
}

export function formatOfExt(ext: string): FormatId | null {
  const e = ext.toLowerCase();
  return OPEN_ORDER.find((f) => FORMATS[f].extensions.includes(e)) ?? null;
}

export const isWritable = (f: FormatId) => FORMATS[f].saveExt !== null;

/** Biçimi içerik imzasıyla (uzantı yanlış ya da eksik olsa da) belirler. */
export async function detectFormat(name: string, data: ArrayBuffer | Uint8Array): Promise<FormatId> {
  const b = data instanceof Uint8Array ? data : new Uint8Array(data);
  const byExt = formatOfExt(extensionOf(name));
  if (isCfb(b)) return "doc";
  if (b[0] === 0x50 && b[1] === 0x4b && b[2] === 0x03 && b[3] === 0x04) {
    try {
      const zip = await JSZip.loadAsync(b);
      const mime = (await zip.file("mimetype")?.async("string"))?.trim() ?? "";
      if (/^application\/vnd\.oasis\.opendocument\.text(-template)?$/.test(mime)) return "odt";
      if (zip.file("[Content_Types].xml")) return "docx";
      if (zip.file("content.xml")) return "odt";
    } catch {
      // Bozuk paket: uzantıya göre denenir.
    }
    return byExt === "odt" ? "odt" : "docx";
  }
  // Metin tabanlı: RTF imzası, yoksa uzantı.
  let i = 0;
  if (b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) i = 3;
  while (i < b.length && i < 64 && (b[i] === 0x20 || b[i] === 0x0a || b[i] === 0x0d || b[i] === 0x09)) i++;
  if (String.fromCharCode(...b.subarray(i, i + 5)) === "{\\rtf") return "rtf";
  const head = new TextDecoder("latin1").decode(b.subarray(i, i + 1024)).toLowerCase();
  // Düz XML OpenDocument metin (.fodt): tek XML belgesi.
  if (/^<(\?xml|office:document)/.test(head)) {
    const top = new TextDecoder("latin1").decode(b.subarray(i, i + 4096));
    if (/<office:document[\s>]/.test(top) && (/office:mimetype="application\/vnd\.oasis\.opendocument\.text(-template)?"/.test(top) || byExt === "fodt")) return "fodt";
  }
  // .doc uzantılı HTML (eski "web sayfası" kayıtları)
  if ((byExt === "doc" || byExt === null) && /^<(!doctype html|html|\?xml)/.test(head) && head.includes("<html")) return "html";
  // İmzası tutmayan ikili/paket uzantıları ilgili okuyucuda anlaşılır hatayla reddedilir.
  return byExt ?? "txt";
}
