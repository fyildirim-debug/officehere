// Dışa aktarılan PDF'in yapısı (sayfa, yer imleri, PDF/A işaretleri, yazı tipleri) ve sayfa metni.
// Metin, Skia'nın yazdığı ToUnicode eşlemlerinden çıkarılır (CI makinesinde poppler gerekmesin).
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";

// pdf-lib masaüstü paketinin bağımlılığıdır; tür tanımı yerine gereken kadar arayüz.
interface Obj {
  toString(): string;
}
interface Dict extends Obj {
  get(k: Obj): Obj | undefined;
  lookup(k: Obj): Obj | undefined;
  has(k: Obj): boolean;
  entries(): [Obj, Obj][];
}
interface Lib {
  PDFDocument: { load(b: Uint8Array, o?: object): Promise<Doc> };
  PDFName: { of(s: string): Obj };
  PDFDict: new (...a: never[]) => Dict;
  PDFArray: new (...a: never[]) => Obj & { asArray(): Obj[] };
  PDFRawStream: new (...a: never[]) => Obj & { dict: Dict };
  decodePDFRawStream(s: Obj): { decode(): Uint8Array };
}
interface Page {
  node: Dict & { Resources(): Dict | undefined; Contents(): Obj | undefined };
}
interface Doc {
  catalog: Dict;
  context: { lookup(o: Obj | undefined): Obj | undefined; trailerInfo: { ID?: Obj }; enumerateIndirectObjects(): [Obj, Obj][] };
  getPages(): Page[];
  getTitle(): string | undefined;
  getAuthor(): string | undefined;
  getCreator(): string | undefined;
}

const lib = createRequire(path.join(__dirname, "..", "apps", "desktop", "package.json"))("@cantoo/pdf-lib") as Lib;
const N = (s: string) => lib.PDFName.of(s);

function textOf(o: Obj | undefined): string {
  const t = o as { decodeText?: () => string } | undefined;
  return t?.decodeText ? t.decodeText() : String(o ?? "");
}

function streamBytes(doc: Doc, o: Obj | undefined): Uint8Array {
  const s = doc.context.lookup(o);
  return s instanceof lib.PDFRawStream ? lib.decodePDFRawStream(s).decode() : new Uint8Array();
}

const utf16 = (hex: string) => {
  let s = "";
  for (let i = 0; i + 4 <= hex.length; i += 4) s += String.fromCharCode(parseInt(hex.slice(i, i + 4), 16));
  return s;
};

/** ToUnicode CMap: bfchar ve bfrange (dizi biçimi dahil). */
function toUnicode(cmap: string): Map<number, string> {
  const m = new Map<number, string>();
  for (const block of cmap.matchAll(/beginbfchar([\s\S]*?)endbfchar/g)) {
    for (const [, src, dst] of block[1].matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g)) m.set(parseInt(src, 16), utf16(dst));
  }
  for (const block of cmap.matchAll(/beginbfrange([\s\S]*?)endbfrange/g)) {
    for (const [, a, b, dst, arr] of block[1].matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(?:<([0-9A-Fa-f]+)>|\[([^\]]*)\])/g)) {
      const lo = parseInt(a, 16);
      const hi = parseInt(b, 16);
      const list = arr ? [...arr.matchAll(/<([0-9A-Fa-f]+)>/g)].map((x) => utf16(x[1])) : null;
      for (let c = lo; c <= hi; c++) {
        if (list) m.set(c, list[c - lo] ?? "");
        else {
          const base = utf16(dst!);
          m.set(c, base.slice(0, -1) + String.fromCharCode(base.charCodeAt(base.length - 1) + (c - lo)));
        }
      }
    }
  }
  return m;
}

function pageText(doc: Doc, page: Page): string {
  const fonts = new Map<string, { map: Map<number, string>; bytes: number }>();
  const fontDict = page.node.Resources()?.lookup(N("Font"));
  if (fontDict instanceof lib.PDFDict) {
    for (const [k, v] of fontDict.entries()) {
      const f = doc.context.lookup(v) as Dict;
      const two = f.get(N("Subtype"))?.toString() === "/Type0";
      const cmap = Buffer.from(streamBytes(doc, f.get(N("ToUnicode")))).toString("latin1");
      fonts.set(k.toString().slice(1), { map: toUnicode(cmap), bytes: two ? 2 : 1 });
    }
  }
  const c = page.node.Contents();
  const parts = c instanceof lib.PDFArray ? c.asArray() : [c];
  const content = parts.map((p) => Buffer.from(streamBytes(doc, p)).toString("latin1")).join("\n");
  let font = fonts.values().next().value;
  let out = "";
  for (const m of content.matchAll(/\/([^\s/[\]<>()]+)\s+[-\d.]+\s+Tf|<([0-9A-Fa-f\s]*)>|\bET\b/g)) {
    if (m[1]) font = fonts.get(m[1]) ?? font;
    else if (m[2] !== undefined && font) {
      const hex = m[2].replace(/\s+/g, "");
      const w = font.bytes * 2;
      for (let i = 0; i + w <= hex.length; i += w) out += font.map.get(parseInt(hex.slice(i, i + w), 16)) ?? "";
    } else out += "\n";
  }
  return out.replace(/[ \t]+/g, " ").trim();
}

export interface PdfBilgi {
  sayfa: number;
  metin: string[];
  yerImleri: string[];
  baslik?: string;
  yazar?: string;
  olusturan?: string;
  /** XMP'deki pdfaid:part / conformance ("2B", "1B"); yoksa null. */
  pdfa: string | null;
  outputIntent: boolean;
  id: boolean;
  etiketli: boolean;
  yaziTipleri: { ad: string; gomulu: boolean; cidSet: boolean }[];
}

export async function pdfBilgi(file: string): Promise<PdfBilgi> {
  const doc = await lib.PDFDocument.load(readFileSync(file), { updateMetadata: false });
  const cat = doc.catalog;
  const yerImleri: string[] = [];
  const walk = (node: Obj | undefined) => {
    for (let it = node instanceof lib.PDFDict ? node.lookup(N("First")) : undefined; it instanceof lib.PDFDict; it = it.lookup(N("Next"))) {
      yerImleri.push(textOf(it.lookup(N("Title"))));
      walk(it);
    }
  };
  walk(cat.lookup(N("Outlines")));
  const xmp = Buffer.from(streamBytes(doc, cat.get(N("Metadata")))).toString("utf8");
  const part = /<pdfaid:part>(\d)<\/pdfaid:part>/.exec(xmp)?.[1];
  const conf = /<pdfaid:conformance>([A-Za-z])<\/pdfaid:conformance>/.exec(xmp)?.[1];
  const yaziTipleri: PdfBilgi["yaziTipleri"] = [];
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof lib.PDFDict) || obj.get(N("Type"))?.toString() !== "/Font") continue;
    const sub = obj.get(N("Subtype"))?.toString();
    if (sub === "/Type0") continue;
    const fd = obj.lookup(N("FontDescriptor"));
    const gomulu = sub === "/Type3" || (fd instanceof lib.PDFDict && ["FontFile", "FontFile2", "FontFile3"].some((k) => fd.has(N(k))));
    yaziTipleri.push({ ad: String(obj.get(N("BaseFont")) ?? sub), gomulu, cidSet: fd instanceof lib.PDFDict && fd.has(N("CIDSet")) });
  }
  return {
    sayfa: doc.getPages().length,
    metin: doc.getPages().map((p) => pageText(doc, p)),
    yerImleri,
    baslik: doc.getTitle(),
    yazar: doc.getAuthor(),
    olusturan: doc.getCreator(),
    pdfa: part ? `${part}${conf ?? ""}` : null,
    outputIntent: !!cat.get(N("OutputIntents")),
    id: !!doc.context.trailerInfo.ID,
    etiketli: !!cat.get(N("StructTreeRoot")),
    yaziTipleri,
  };
}
