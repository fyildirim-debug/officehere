/**
 * XLSX paketini okumadan önce denetler ve onarır (exceljs'e verilmeden):
 * - Paket sınırları: parça sayısı, bildirilen açılmış boyut toplamı (sıkıştırma
 *   bombası), tutarsız boyut bilgisi; aşırı sıkıştırılmış XML'deki dev boşluk
 *   dolguları ayıklanır.
 * - Eksik/bozuk yapı: workbook.xml, ilişkiler, stiller ya da sayfa parçası
 *   yoksa kitap sessizce boş açılmaz: sayfalar kurtarılır ve uyarı verilir
 *   (loadWarnings dolu olduğundan otomatik kayıt özgünün üzerine yazmaz).
 * - Sayfa XML'i: yarıda kesilmiş sayfa son tam satıra kadar kurtarılır; Excel
 *   sınırları dışındaki satır/hücre numaraları düzeltilir; sütun aralıkları
 *   kırpılır; geçersiz ve çakışan birleştirmeler ayıklanır; ilişkisi olmayan
 *   çizim/köprü/tablo başvuruları çıkarılır.
 * - exceljs'e verilen kopyada tablo parçaları zararsız bir taslakla değiştirilir
 *   (tablolar xlsxTables ile doğrudan okunur) ve birleştirmeler tümüyle çıkarılır:
 *   exceljs her birleştirmede öncekilerin hepsini gezer (50.000 birleştirmede
 *   dakikalar) ve birleştirmedeki her hücre için nesne üretir (A1:XFD1048576 =
 *   17 milyar hücre). Birleştirmeler buradaki taramayla okunur (merges).
 */
import JSZip from "jszip";
import { colToName, MAX_COLS, MAX_ROWS, nameToCol, parseRangeAddress } from "@officehere/formula";
import type { Rect } from "../model/types";
import { disjointMerges, FormatError } from "./bookUtil";
import { toOpenError } from "./openErrors";
import { relsPathOf, unescapeXml } from "./ooxmlZip";
import { normalizeZip, resolvePath } from "./xlsxPackage";
import { memoizeText } from "./zipCache";

const MB = 1024 * 1024;

export const XLSX_LIMITS = {
  /** Paketteki en çok parça sayısı. */
  entries: 20_000,
  /** Bildirilen açılmış boyutların toplamı. */
  total: 1024 * MB,
  /** Tek bir XML parçası (tarayıcıdaki en uzun metin sınırının altında). */
  part: 400 * MB,
  /** Bu boyutu ve sıkıştırma oranını aşan parçalarda dev boşluk dolguları ayıklanır. */
  bombPart: 8 * MB,
  bombRatio: 200,
};

const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_TYPE = NS_REL + "/";

export interface PreparedXlsx {
  /** Onarılmış paket: kendi okuyucularımız için. */
  data: ArrayBuffer;
  /** exceljs'e verilecek paket (tablo parçaları taslakla değiştirilmiş, birleştirmeler çıkarılmış). */
  excel: ArrayBuffer;
  zip: JSZip;
  /** Sayfa adı → birleştirmeler (geçerli, çakışmasız, dosyadaki sırayla; exceljs'e verilmez). */
  merges: Map<string, Rect[]>;
  /** Okunamayan, atlanan ya da onarılan öğeler (loadWarnings). */
  warnings: string[];
  /** Yapısal onarım yapıldıysa kullanıcıya gösterilecek ileti. */
  notice?: string;
  /** Sayfa XML'lerinde hücre verisi vardı (okuma sonrası hiç hücre yoksa sessiz kayıp sayılır). */
  hadCells: boolean;
  /** Paketteki (onarılmış) çalışma sayfalarının adları. */
  sheetNames: string[];
}

interface Rel {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

function attrOf(tag: string, name: string): string | undefined {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag);
  return m ? m[1] : undefined;
}

function setAttr(attrs: string, name: string, value: string): string {
  const re = new RegExp(`(\\s${name}=")[^"]*(")`);
  return re.test(attrs) ? attrs.replace(re, `$1${value}$2`) : ` ${name}="${value}"${attrs}`;
}

function parseRels(xml: string): Rel[] {
  const out: Rel[] = [];
  for (const m of xml.matchAll(/<(?:\w+:)?Relationship\b[^>]*>/g)) {
    const id = attrOf(m[0], "Id");
    const type = attrOf(m[0], "Type") ?? "";
    const target = attrOf(m[0], "Target");
    if (id && target !== undefined) out.push({ id, type, target: unescapeXml(target), external: attrOf(m[0], "TargetMode") === "External" });
  }
  return out;
}

const relKind = (type: string) => type.slice(type.lastIndexOf("/") + 1);

/** Öznitelikteki sayfa adı: adlandırılmış ve sayısal karakter başvuruları çözülür. */
const decodeName = (v: string) => unescapeXml(v.replace(/&#(x[0-9a-fA-F]+|\d+);/g, (_m, n: string) => String.fromCodePoint(n[0] === "x" ? parseInt(n.slice(1), 16) : Number(n))));

/** Doğal sıralama: sheet2 < sheet10. */
function natural(a: string, b: string): number {
  return a.localeCompare(b, undefined, { numeric: true });
}

/** Eksik parça türü → kullanıcıya gösterilecek ad. */
const LOST_LABEL: Record<string, string> = {
  comments: "açıklamalar (notlar)",
  drawing: "resimler, grafikler ve çizimler",
  chart: "grafikler",
  image: "resimler",
  table: "tablolar",
  pivotTable: "özet tablolar",
  pivotCacheDefinition: "özet tablolar",
  hyperlink: "köprüler",
  legacyDrawingHF: "üst/alt bilgi resimleri",
  picture: "sayfa arka planı",
  tablePart: "tablolar",
};

const declared = (f: JSZip.JSZipObject) => (f as unknown as { _data?: { uncompressedSize?: number; compressedSize?: number } })._data;

const EMPTY_SHEET = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><sheetData/></worksheet>`;
const minStyles = (dxfs: number) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${NS_MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
  `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
  `<dxfs count="${Math.max(1, dxfs)}">${"<dxf/>".repeat(Math.max(1, dxfs))}</dxfs></styleSheet>`;
const emptySst = (n: number) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS_MAIN}" count="${n}" uniqueCount="${n}">${"<si><t></t></si>".repeat(n)}</sst>`;

/**
 * DOCTYPE bildirimi (OOXML'de yasak; "milyar kahkaha" varlık genişlemesi) ve tanımsız
 * varlık başvuruları çıkarılır; bildirim yoksa null.
 */
export function stripDoctype(xml: string): string | null {
  const head = xml.slice(0, 65536);
  const at = head.search(/<!DOCTYPE\b/i);
  if (at < 0) {
    // Bildirim yok ama tanımsız varlık başvurusu var (ör. &lol;): exceljs "undefined entity" ile düşer
    return BAD_ENTITY.test(xml) ? xml.replace(BAD_ENTITY_G, "") : null;
  }
  // İç alt küme ([...]) dahil bildirimin sonu
  const sub = head.indexOf("[", at);
  const gt = head.indexOf(">", at);
  let end = gt;
  if (sub >= 0 && sub < gt) {
    const close = xml.indexOf("]>", sub);
    end = close >= 0 ? close + 1 : gt;
  }
  const body = xml.slice(0, at) + xml.slice(end + 1);
  return body.replace(BAD_ENTITY_G, "");
}

const BAD_ENTITY = /&(?!(?:lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);)[A-Za-z_][\w.-]*;/;
const BAD_ENTITY_G = new RegExp(BAD_ENTITY.source, "g");
/** exceljs'e tablo yerine verilen zararsız taslak (tablolar xlsxTables ile okunur). */
const TABLE_STUB = (n: number) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="${NS_MAIN}" id="${n}" name="OhTaslak${n}" displayName="OhTaslak${n}" ref="A1:A2"><tableColumns count="1"><tableColumn id="1" name="A"/></tableColumns></table>`;

/** Paket sınırları ve tutarlılık denetimi; aşırı sıkıştırılmış (bomba şüphesi) parçaları döndürür. */
export function checkPackageLimits(zip: JSZip): Set<string> {
  const files = Object.values(zip.files).filter((f) => !f.dir);
  if (files.length > XLSX_LIMITS.entries) throw new FormatError("Dosya olağan dışı sayıda parça içeriyor; güvenlik için açılmadı.", `${files.length} parça`);
  let total = 0;
  const suspicious = new Set<string>();
  for (const f of files) {
    const d = declared(f);
    const u = d?.uncompressedSize;
    const c = d?.compressedSize;
    if (typeof u !== "number" || typeof c !== "number") continue;
    // DEFLATE veriyi en çok birkaç bayt büyütür: açılmış boyut sıkıştırılmıştan çok küçükse bilgi sahte/bozuk
    if (c > 4096 && u + 1024 < c * 0.9) throw new FormatError("Dosya bozuk ya da eksik indirilmiş: paket içindeki boyut bilgileri tutarsız.", `${f.name}: açılmış ${u} B, sıkıştırılmış ${c} B`);
    total += u;
    if (u > XLSX_LIMITS.bombPart && u / Math.max(1, c) > XLSX_LIMITS.bombRatio) suspicious.add(f.name);
    if (u > XLSX_LIMITS.part && !suspicious.has(f.name))
      throw new FormatError(`Dosya bu uygulamada açılamayacak kadar büyük: tek bir bölümü açıldığında ${Math.round(u / MB)} MB tutuyor.`, `${f.name}: ${u} B`);
    if (u > 2 * XLSX_LIMITS.part) throw new FormatError("Dosya olağan dışı ölçüde sıkıştırılmış veri içeriyor (sıkıştırma bombası olabilir); güvenlik için açılmadı.", `${f.name}: ${u} B / ${c} B`);
  }
  if (total > XLSX_LIMITS.total)
    throw new FormatError(`Dosya açılamayacak kadar büyük: içindeki veriler açıldığında ${Math.round(total / MB).toLocaleString("tr-TR")} MB tutuyor (sınır ${XLSX_LIMITS.total / MB} MB).`, `toplam ${total} B`);
  return suspicious;
}

const isWs = (c: number) => c === 32 || c === 10 || c === 13 || c === 9;
const TEXT_CLOSE = /<\/(?:\w+:)?(?:t|v|f)>/y;
const NON_WS = /[^ \t\r\n]/g;

/**
 * Etiketler arasındaki dev boşluk dizilerini (1024+ karakter) parça parça ayıklar.
 * Düzenli ifade geri izlemesi kullanılmaz (yüz milyonlarca karakterde yığın taşar);
 * dizinin sonu yerel aramayla bulunur. Metin içeriği (<t>   </t>) korunur.
 */
export class WsCollapser {
  private parts: string[] = [];
  /** Önceki parça ">" + boşlukla bitti; dizi bu parçada sürüyor. */
  private inRun = false;
  private runLen = 0;
  private runHead = "";

  push(s: string): void {
    const n = s.length;
    let start = 0;
    if (this.inRun) {
      NON_WS.lastIndex = 0;
      const m = NON_WS.exec(s);
      const j = m ? m.index : n;
      if (this.runHead.length < 1024) this.runHead += s.slice(0, Math.min(j, 1024 - this.runHead.length));
      this.runLen += j;
      if (!m) return;
      TEXT_CLOSE.lastIndex = j;
      if (this.runLen < 1024 || TEXT_CLOSE.test(s)) this.parts.push(this.runHead);
      this.inRun = false;
      start = j;
    }
    let last = start;
    let i = start;
    while ((i = s.indexOf(">", i)) >= 0) {
      let j = i + 1;
      if (j < n && isWs(s.charCodeAt(j))) {
        NON_WS.lastIndex = j;
        const m = NON_WS.exec(s);
        j = m ? m.index : n;
      }
      if (j >= n) {
        // Dizi parçanın sonuna dayandı: sonraki parçada sürer
        this.parts.push(s.slice(last, i + 1));
        this.inRun = true;
        this.runLen = n - i - 1;
        this.runHead = s.slice(i + 1, Math.min(n, i + 1025));
        return;
      }
      if (j - i - 1 >= 1024) {
        TEXT_CLOSE.lastIndex = j;
        if (!TEXT_CLOSE.test(s)) {
          this.parts.push(s.slice(last, i + 1));
          last = j;
        }
      }
      i = j;
    }
    this.parts.push(s.slice(last));
  }

  end(): string {
    if (this.inRun && this.runLen < 1024) this.parts.push(this.runHead);
    this.inRun = false;
    return this.parts.join("");
  }
}

/** Aşırı sıkıştırılmış XML'deki dev boşluk dolgularını (etiketler arası, 1024+ karakter) ayıklar. */
export function collapseWhitespace(xml: string): string {
  const c = new WsCollapser();
  c.push(xml);
  return c.end();
}

/**
 * Bomba şüphesi olan parçayı akışla açıp ayıklar: tarayıcının yerel açıcısı
 * (DecompressionStream) JSZip'ten kat kat hızlıdır ve yüz milyonlarca karakterlik
 * ara metin üretilmez. Yoksa JSZip ile açılıp ayıklanır.
 */
export async function readCollapsed(f: JSZip.JSZipObject): Promise<string> {
  const d = (f as unknown as { _data?: { compressedContent?: Uint8Array; compression?: { magic?: string } } })._data;
  const DS = (globalThis as unknown as { DecompressionStream?: new (format: string) => TransformStream<Uint8Array, Uint8Array> }).DecompressionStream;
  if (!DS || !d?.compressedContent || d.compression?.magic !== "\x08\x00" || typeof Blob === "undefined") return collapseWhitespace(await f.async("string"));
  const reader = new Blob([d.compressedContent as BlobPart]).stream().pipeThrough(new DS("deflate-raw")).getReader();
  const dec = new TextDecoder("utf-8");
  const c = new WsCollapser();
  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    c.push(dec.decode(value, { stream: true }));
  }
  c.push(dec.decode());
  return c.end();
}

/** Yarıda kesilmiş sayfa XML'ini son tam satıra kadar kurtarır; sağlamsa null. */
export function repairTruncatedSheet(xml: string): string | null {
  const tail = xml.slice(-256).trimEnd();
  if (/<\/(?:\w+:)?worksheet>$/.test(tail)) return null;
  const root = /<((?:\w+:)?)worksheet\b[^>]*>/.exec(xml);
  if (!root) return EMPTY_SHEET;
  const p = root[1];
  const sd = new RegExp(`<${p}sheetData\\b[^>]*?(/?)>`).exec(xml);
  if (!sd) return `${xml.slice(0, root.index + root[0].length)}<${p}sheetData/></${p}worksheet>`;
  if (sd[1]) return `${xml.slice(0, sd.index + sd[0].length)}</${p}worksheet>`;
  const end = xml.indexOf(`</${p}sheetData>`);
  if (end >= 0) {
    // sheetData tam; sonrası kesik: sheetData'dan sonrası atılır
    return `${xml.slice(0, end)}</${p}sheetData></${p}worksheet>`;
  }
  const body = xml.slice(sd.index + sd[0].length);
  let cut = -1;
  for (const m of body.matchAll(new RegExp(`</${p}row>|<${p}row\\b[^>]*/>`, "g"))) cut = m.index! + m[0].length;
  const rows = cut >= 0 ? body.slice(0, cut) : "";
  return `${xml.slice(0, sd.index + sd[0].length)}${rows}</${p}sheetData></${p}worksheet>`;
}

const ROW_TAG = /<((?:\w+:)?)row\b([^>]*?)(\/>|>([\s\S]*?)<\/\1row>)/g;
const CELL_TAG = /<((?:\w+:)?)c\b([^>]*?)(\/>|>([\s\S]*?)<\/\1c>)/g;

/**
 * Satır/hücre numaralarını denetler: geçersiz (eksik, 0, eksi, sayı değil, Excel
 * sınırı dışında) olanlar sıradaki numaraya çekilir; sınır dışına taşanlar atılır.
 * Yalnız sorun varsa yeniden yazar.
 */
export function repairRowNumbers(xml: string): { xml: string; dropped: number } | null {
  const sdOpen = /<(?:\w+:)?sheetData\b[^>]*?(\/?)>/.exec(xml);
  if (!sdOpen || sdOpen[1]) return null;
  const start = sdOpen.index + sdOpen[0].length;
  const end = xml.indexOf("sheetData>", start);
  if (end < 0) return null;
  const close = xml.lastIndexOf("<", end);
  const body = xml.slice(start, close);
  let bad = false;
  for (const m of body.matchAll(/<(?:\w+:)?row\b([^>]*)>/g)) {
    const r = attrOf(m[1], "r");
    if (r === undefined || !/^\d{1,7}$/.test(r) || +r < 1 || +r > MAX_ROWS) {
      bad = true;
      break;
    }
  }
  // Hücre başvurusu biçimi: 1-3 harf + satır; harfsiz, eksik ya da XFD'yi aşan sütun
  if (!bad && /<(?:\w+:)?c(?=[\s/>])(?![^>]*\sr="[A-Za-z]{1,3}[1-9]\d{0,6}")[^>]*>/.test(body)) bad = true;
  if (!bad && /<(?:\w+:)?c(?=\s)[^>]*\sr="(?:X(?:F[E-Z]|[G-Z][A-Z])|[Y-Z][A-Z]{2}|[A-Z]{4,})\d/i.test(body)) bad = true;
  if (!bad) return null;
  let prevRow = 0;
  let dropped = 0;
  const fixed = body.replace(ROW_TAG, (_m, p: string, attrs: string, _end: string, inner: string | undefined) => {
    const rs = attrOf(attrs, "r");
    let r = rs !== undefined && /^\d{1,7}$/.test(rs) ? Number(rs) : NaN;
    if (!(r >= 1 && r <= MAX_ROWS)) r = prevRow + 1;
    if (r > MAX_ROWS) {
      dropped++;
      return "";
    }
    prevRow = Math.max(prevRow, r);
    const a = setAttr(attrs, "r", String(r));
    if (inner === undefined) return `<${p}row${a}/>`;
    let prevCol = 0;
    const cells = inner.replace(CELL_TAG, (cm: string, cp: string, cattrs: string) => {
      const ref = attrOf(cattrs, "r");
      const letters = ref ? /^([A-Za-z]{1,3})\d*$/.exec(ref)?.[1] : undefined;
      let col = letters ? nameToCol(letters) + 1 : 0;
      if (col < 1) col = prevCol + 1;
      if (col > MAX_COLS) {
        dropped++;
        return "";
      }
      prevCol = col;
      return `<${cp}c${setAttr(cattrs, "r", `${colToName(col - 1)}${r}`)}${cm.slice(cp.length + 2 + cattrs.length)}`;
    });
    return `<${p}row${a}>${cells}</${p}row>`;
  });
  return { xml: xml.slice(0, start) + fixed + xml.slice(close), dropped };
}

/** <col min max>: Excel sınırına kırpılır (exceljs aralıktaki her sütun için nesne üretir). */
export function repairCols(xml: string): string | null {
  const sd = xml.search(/<(?:\w+:)?sheetData\b/);
  const head = sd >= 0 ? xml.slice(0, sd) : xml;
  if (!/<(?:\w+:)?col\b/.test(head)) return null;
  let changed = false;
  const fixed = head.replace(/<((?:\w+:)?)col\b([^>]*?)\/>/g, (m, p: string, attrs: string) => {
    let min = Number(attrOf(attrs, "min"));
    let max = Number(attrOf(attrs, "max"));
    if (Number.isInteger(min) && Number.isInteger(max) && min >= 1 && max >= min && max <= MAX_COLS) return m;
    changed = true;
    if (!Number.isFinite(min) || min < 1) min = 1;
    if (!Number.isFinite(max) || max < min) max = min;
    if (min > MAX_COLS) return "";
    max = Math.min(max, MAX_COLS);
    return `<${p}col${setAttr(setAttr(attrs, "min", String(Math.floor(min))), "max", String(Math.floor(max)))}/>`;
  });
  return changed ? fixed + xml.slice(head.length) : null;
}

/**
 * Birleştirmeler: ayrıştırılamayan, tek hücrelik ve çakışan aralıklar atılır (blok yalnız o zaman
 * yeniden yazılır); geçerli liste dosyadaki sırayla döner. `excel`: mergeCells'i tümüyle çıkarılmış
 * XML (exceljs'e verilecek kopya). Kapanmamış blok son birleştirme öğesine kadar okunur; birden çok
 * blok tek blokta toplanır. `big`: küçük birleştirmeyle çakıştığı için atılan dev birleştirmeler.
 * Tarama doğrusal, çakışma denetimi O(n log n).
 */
export function repairMerges(xml: string): { xml: string; excel: string; merges: Rect[]; dropped: number; big: number } | null {
  const spans: [number, number][] = [];
  const all: Rect[] = [];
  let bad = 0;
  let p = "";
  const OPEN = /<((?:\w+:)?)mergeCells\b[^>]*?(\/?)>/g;
  for (let open = OPEN.exec(xml); open; open = OPEN.exec(xml)) {
    const at = open.index;
    const from = at + open[0].length;
    let end = from;
    if (!open[2]) {
      const closeTag = `</${open[1]}mergeCells>`;
      const close = xml.indexOf(closeTag, from);
      if (close >= 0) end = close + closeTag.length;
      else {
        bad++;
        const item = /<(?:\w+:)?mergeCell\b[^>]*>(?:\s*<\/(?:\w+:)?mergeCell>)?/g;
        item.lastIndex = from;
        while (item.exec(xml)) end = item.lastIndex;
      }
    }
    for (const x of xml.slice(from, end).matchAll(/<(?:\w+:)?mergeCell\b[^>]*>/g)) {
      const r = parseRangeAddress((/\sref="([^"]*)"/.exec(x[0])?.[1] ?? "").replace(/\$/g, ""));
      if (r) all.push(r);
      else bad++;
    }
    if (!spans.length) p = open[1];
    spans.push([at, end]);
    OPEN.lastIndex = end;
  }
  if (!spans.length) return null;
  // Tek hücrelik birleştirme anlamsızdır, sessizce atılır (invalid); ayrıştırılamayan ve çakışanlar uyarılır
  const { merges, invalid, overlapping, bigDropped } = disjointMerges(all);
  const splice = (first: string) => {
    let out = "";
    let k = 0;
    spans.forEach(([a, b], j) => {
      out += xml.slice(k, a) + (j ? "" : first);
      k = b;
    });
    return out + xml.slice(k);
  };
  const excel = splice("");
  if (!bad && !invalid && !overlapping && !bigDropped && spans.length === 1) return { xml, excel, merges, dropped: 0, big: 0 };
  const ref = (r: Rect) => `${colToName(r.c1)}${r.r1 + 1}:${colToName(r.c2)}${r.r2 + 1}`;
  const block = merges.length ? `<${p}mergeCells count="${merges.length}">${merges.map((r) => `<${p}mergeCell ref="${ref(r)}"/>`).join("")}</${p}mergeCells>` : "";
  return { xml: splice(block), excel, merges, dropped: bad + overlapping, big: bigDropped };
}

/** İlişkisi olmayan r:id başvuruları (silinmiş .rels ya da eksik giriş): exceljs "Target" hatasıyla düşer. */
export function dropDanglingRefs(xml: string, rels: Rel[] | null, exists: (r: Rel) => boolean): { xml: string; lost: Set<string> } | null {
  const ids = new Map((rels ?? []).map((r) => [r.id, r]));
  const dangling = (id: string) => {
    const r = ids.get(id);
    return !r || (!r.external && relKind(r.type) !== "hyperlink" && !exists(r));
  };
  if (!xml.includes(':id="')) return null;
  const lost = new Set<string>();
  let out = xml.replace(/<((?:\w+:)?)(drawing|legacyDrawing|legacyDrawingHF|picture|tablePart)\b[^>]*?\s\w+:id="([^"]*)"[^>]*\/>/g, (m, _p: string, el: string, id: string) => {
    if (!dangling(id)) return m;
    // VML not çizimi yalnız görünümdür: notlar comments parçasından okunur
    if (el !== "legacyDrawing") lost.add(LOST_LABEL[el] ?? "diğer nesneler");
    return "";
  });
  out = out.replace(/<((?:\w+:)?)hyperlink\b([^>]*?)\s\w+:id="([^"]*)"([^>]*?)(\/>|>[\s\S]*?<\/\1hyperlink>)/g, (m, p: string, a: string, id: string, b: string) => {
    const r = ids.get(id);
    if (r) return m;
    lost.add("köprüler");
    // Konumlu (iç) köprü korunur, yalnız ilişki başvurusu kalkar
    return /\slocation="/.test(a + b) ? `<${p}hyperlink${a}${b}/>` : "";
  });
  out = out.replace(/<((?:\w+:)?)(oleObjects|controls)\b[\s\S]*?<\/\1\2>/g, (m) => {
    const bad = [...m.matchAll(/\s\w+:id="([^"]*)"/g)].some((x) => dangling(x[1]));
    if (!bad) return m;
    lost.add("gömülü nesneler ve denetimler");
    return "";
  });
  if (out === xml) return null;
  // Boş kalan ya da sayısı değişen tableParts
  out = out.replace(/<((?:\w+:)?)tableParts\b[^>]*?(?:\/>|>\s*<\/\1tableParts>)/g, "");
  out = out.replace(/<((?:\w+:)?)tableParts\b([^>]*)>([\s\S]*?)<\/\1tableParts>/g, (_m, p: string, a: string, inner: string) => {
    const n = (inner.match(/<(?:\w+:)?tablePart\b/g) ?? []).length;
    return `<${p}tableParts${setAttr(a, "count", String(n))}>${inner}</${p}tableParts>`;
  });
  return { xml: out, lost };
}

/** Paket onarıcı: yapı ve sayfa denetimleri. */
class Guard {
  warnings: string[] = [];
  notices: string[] = [];
  merges = new Map<string, Rect[]>();
  /** Parça yolu → exceljs'e verilecek (birleştirmeleri çıkarılmış) sayfa XML'i. */
  excelSheets = new Map<string, string>();
  hadCells = false;
  /** Sayfa adı → parça yolu. */
  sheets: { name: string; path: string }[] = [];
  changed = false;

  constructor(
    readonly zip: JSZip,
    readonly suspicious: Set<string>,
  ) {}

  text(path: string): Promise<string | undefined> {
    return this.zip.file(path)?.async("string") ?? Promise.resolve(undefined);
  }

  put(path: string, content: string, ct?: string): void {
    this.zip.file(path, content, { compression: content.length > 4 * MB ? "STORE" : "DEFLATE" });
    this.changed = true;
    if (ct) this.pendingCt.push([path, ct]);
  }

  private pendingCt: [string, string][] = [];

  warn(label: string): void {
    if (!this.warnings.includes(label)) this.warnings.push(label);
  }

  /** workbook.xml ve ilişkileri: eksikse kurtarılır, sayfa→parça eşlemesi çıkarılır. */
  async repairWorkbook(): Promise<void> {
    const zip = this.zip;
    const sheetFiles = Object.keys(zip.files)
      .filter((n) => /^xl\/worksheets\/[^/]+\.xml$/i.test(n) && !zip.files[n].dir)
      .sort(natural);
    let wbXml = await this.text("xl/workbook.xml");
    const wellFormed = !!wbXml && /<\/(?:\w+:)?workbook>\s*$/.test(wbXml.slice(-256));
    const entries = wbXml ? [...wbXml.matchAll(/<(?:\w+:)?sheet\b[^>]*?\/?>/g)].map((m) => ({ tag: m[0], name: decodeName(attrOf(m[0], "name") ?? ""), rid: /\s\w+:id="([^"]*)"/.exec(m[0])?.[1] })) : [];
    if (!entries.length) {
      if (!sheetFiles.length) throw new FormatError("Dosya bozuk ya da eksik: çalışma kitabında okunabilir bir sayfa bulunamadı.", wbXml ? "workbook.xml: <sheet> yok" : "xl/workbook.xml yok");
      // Çalışma kitabı yapısı yok: sayfa parçalarından yeni yapı kurulur
      this.sheets = sheetFiles.map((path, i) => ({ name: `Sayfa${i + 1}`, path }));
      this.writeWorkbook(this.sheets.map((s, i) => ({ name: s.name, rid: `rIdOh${i + 1}` })), "");
      this.writeWorkbookRels(this.sheets.map((s, i) => ({ id: `rIdOh${i + 1}`, path: s.path })), []);
      this.warn("sayfa adları ve çalışma kitabı ayarları (dosyanın kitap yapısı eksikti)");
      this.notices.push(`Dosya bozuktu: çalışma kitabı yapısı eksikti; ${this.sheets.length} sayfa kurtarıldı. Sayfa adları ve bazı ayarlar kaybolmuş olabilir; kaydetmeden önce içeriği denetleyin.`);
      return;
    }
    const relsXml = await this.text("xl/_rels/workbook.xml.rels");
    const rels = relsXml ? parseRels(relsXml) : [];
    const byId = new Map(rels.map((r) => [r.id, r]));
    const used = new Set<string>();
    const resolved: { name: string; path?: string; rid?: string; kind?: string }[] = [];
    for (const e of entries) {
      const r = e.rid ? byId.get(e.rid) : undefined;
      const path = r && !r.external ? resolvePath("xl/workbook.xml", r.target) : undefined;
      // Hedef gerçekten bir sayfa parçası olmalı (döngüsel ilişki workbook.xml'i gösterebilir) ve iki sayfa aynı parçayı paylaşamaz
      const kind = r ? relKind(r.type) : undefined;
      if (r && path && zip.file(path) && !used.has(path) && (kind !== "worksheet" || (await this.isWorksheet(path)))) {
        used.add(path);
        resolved.push({ name: e.name, path, rid: e.rid, kind });
      } else resolved.push({ name: e.name, kind });
    }
    const missing = resolved.filter((s) => !s.path && s.kind !== "chartsheet" && s.kind !== "dialogsheet" && s.kind !== "macrosheet");
    let relsChanged = !relsXml;
    if (missing.length) {
      // İlişkisi kopuk sayfalar: kullanılmayan sayfa parçalarıyla sırayla eşleştirilir; kalmazsa boş sayfa
      const spare: string[] = [];
      for (const p of sheetFiles) if (!used.has(p) && (await this.isWorksheet(p))) spare.push(p);
      let recovered = 0;
      let emptied = 0;
      for (const s of missing) {
        const path = spare.shift();
        if (path) {
          s.path = path;
          recovered++;
        } else {
          let n = 1;
          while (zip.file(`xl/worksheets/sheet${n}.xml`)) n++;
          s.path = `xl/worksheets/sheet${n}.xml`;
          this.put(s.path, EMPTY_SHEET, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
          this.warn(`"${s.name}" sayfasının içeriği (dosyada yok)`);
          emptied++;
        }
        s.kind = "worksheet";
        s.rid = undefined;
      }
      relsChanged = true;
      if (recovered) this.notices.push(`Dosya bozuktu: sayfa bağlantıları eksikti; ${recovered} sayfa sırayla eşleştirilerek kurtarıldı. Kaydetmeden önce sayfaların içeriğini denetleyin.`);
      if (recovered) this.warn("sayfa bağlantıları (sayfalar sırayla eşleştirildi)");
      if (emptied) this.notices.push(`Dosya bozuktu: ${emptied} sayfanın içeriği dosyada yok; bu sayfalar boş açıldı.`);
    }
    let n = 1;
    const ridFor = new Map<string, string>();
    for (const s of resolved) {
      if (!s.path) continue;
      if (!s.rid) {
        while (byId.has(`rIdOh${n}`)) n++;
        s.rid = `rIdOh${n++}`;
        ridFor.set(s.rid, s.path);
      }
    }
    if (relsChanged || !wellFormed) {
      // İlişkiler yeniden yazılır: geçerli eski girişler + yeni sayfa girişleri; kopuk iç hedefler atılır
      const keep = rels.filter((r) => r.external || !!zip.file(resolvePath("xl/workbook.xml", r.target)) || relKind(r.type) === "worksheet");
      const live = keep.filter((r) => !(relKind(r.type) === "worksheet" && !resolved.some((s) => s.rid === r.id)));
      this.writeWorkbookRels(
        [...ridFor].map(([id, path]) => ({ id, path })),
        live,
      );
    }
    this.sheets = resolved.filter((s) => s.path && s.kind === "worksheet").map((s) => ({ name: s.name, path: s.path! }));
    if (!wellFormed) {
      // Bozuk/kesik workbook.xml: sayfalar ve okunabilen tanımlı adlarla yeniden kurulur
      const names = wbXml ? [...wbXml.matchAll(/<(?:\w+:)?definedName\b[^>]*>[^<]*<\/(?:\w+:)?definedName>/g)].map((m) => m[0]).join("") : "";
      this.writeWorkbook(
        resolved.filter((s) => s.path).map((s) => ({ name: s.name, rid: s.rid! })),
        names,
      );
      this.warn("çalışma kitabı ayarları (kitap yapısı bozuktu)");
      this.notices.push("Dosya bozuktu: çalışma kitabı yapısı onarılarak açıldı. Kaydetmeden önce içeriği denetleyin.");
    } else if (ridFor.size && wbXml) {
      // Yeni ilişki numaraları workbook.xml'e işlenir
      let i = 0;
      wbXml = wbXml.replace(/<(?:\w+:)?sheet\b[^>]*?\/?>/g, (tag) => {
        const s = resolved[i++];
        if (!s?.rid) return tag;
        return /\s\w+:id="/.test(tag) ? tag.replace(/(\s\w+:id=")[^"]*(")/, `$1${s.rid}$2`) : tag.replace(/\s*\/?>$/, (end) => ` r:id="${s.rid}"${end.trim()}`);
      });
      if (!/xmlns:r="/.test(wbXml)) wbXml = wbXml.replace(/<((?:\w+:)?)workbook\b/, `<$1workbook xmlns:r="${NS_REL}"`);
      this.put("xl/workbook.xml", wbXml);
    }
  }

  /** Parçanın kök öğesi worksheet mi (boş ya da kesik içerik de sayfa sayılır: onarılır). */
  private async isWorksheet(path: string): Promise<boolean> {
    if (/^xl\/workbook\.xml$|\.rels$|^\[Content_Types\]\.xml$/.test(path)) return false;
    let xml: string | undefined;
    try {
      xml = await this.text(path);
    } catch {
      return true; // açılamayan veri: repairSheets boş sayfayla uyarır
    }
    const root = /<(?:\w+:)?([A-Za-z][\w.-]*)[\s>/]/.exec((xml ?? "").slice(0, 8192).replace(/<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<!DOCTYPE[^>]*>/g, ""));
    return !root || root[1] === "worksheet";
  }

  private writeWorkbook(sheets: { name: string; rid: string }[], names: string): void {
    const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
    const xml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><bookViews><workbookView/></bookViews><sheets>` +
      sheets.map((s, i) => `<sheet name="${esc(s.name || `Sayfa${i + 1}`)}" sheetId="${i + 1}" r:id="${s.rid}"/>`).join("") +
      `</sheets>${names ? `<definedNames>${names}</definedNames>` : ""}</workbook>`;
    this.put("xl/workbook.xml", xml, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
  }

  private writeWorkbookRels(added: { id: string; path: string }[], keep: Rel[]): void {
    const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;");
    const items = keep.map((r) => `<Relationship Id="${esc(r.id)}" Type="${esc(r.type)}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ""}/>`);
    for (const a of added) items.push(`<Relationship Id="${a.id}" Type="${REL_TYPE}worksheet" Target="${a.path.replace(/^xl\//, "")}"/>`);
    const has = (kind: string) => keep.some((r) => relKind(r.type) === kind);
    let k = 1;
    const extra = (kind: string, target: string) => {
      if (has(kind) || !this.zip.file(`xl/${target}`)) return;
      items.push(`<Relationship Id="rIdOhX${k++}" Type="${REL_TYPE}${kind}" Target="${target}"/>`);
    };
    extra("styles", "styles.xml");
    extra("sharedStrings", "sharedStrings.xml");
    extra("theme", "theme/theme1.xml");
    this.put("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.join("")}</Relationships>`);
  }

  /** Sayfa parçalarını denetler/onarır. */
  async repairSheets(): Promise<void> {
    const hasSst = !!this.zip.file("xl/sharedStrings.xml");
    const hasStyles = !!this.zip.file("xl/styles.xml");
    let maxSst = -1;
    let maxDxf = -1;
    for (const s of this.sheets) {
      const f = this.zip.file(s.path);
      if (!f) continue;
      let xml: string;
      try {
        xml = await f.async("string");
      } catch {
        // Tek sayfanın verisi açılamıyor (bozuk sıkıştırma): sayfa boş açılır, uyarılır
        this.put(s.path, EMPTY_SHEET);
        this.warn(`"${s.name}" sayfasının içeriği (bozuk veri)`);
        this.notices.push(`Dosya bozuk: "${s.name}" sayfasının verisi okunamadı ve sayfa boş açıldı.`);
        continue;
      }
      const orig = xml;
      if (!xml.trim()) {
        xml = EMPTY_SHEET;
        this.warn(`"${s.name}" sayfasının içeriği (boş ya da bozuk)`);
        this.notices.push(`Dosya bozuk: "${s.name}" sayfasının verisi boş; sayfa boş açıldı.`);
      }
      const doc = stripDoctype(xml);
      if (doc !== null) {
        xml = doc;
        this.warn("tanımsız XML varlıkları (dosya değiştirilmiş olabilir)");
      }
      const cut = repairTruncatedSheet(xml);
      if (cut !== null) {
        xml = cut;
        this.warn(`"${s.name}" sayfasının bir bölümü (dosya yarıda kesilmiş)`);
        this.notices.push(`Dosya bozuk ya da yarıda kesilmiş: "${s.name}" sayfasının okunabilen bölümü açıldı. Kaydetmeden önce içeriği denetleyin.`);
      }
      const rows = repairRowNumbers(xml);
      if (rows) {
        xml = rows.xml;
        this.warn(rows.dropped ? "sınır dışındaki satır ve hücreler" : "satır/hücre numaraları (onarıldı)");
      }
      const cols = repairCols(xml);
      if (cols !== null) xml = cols;
      // İlişkiler: hedefi olmayan iç ilişkiler çıkarılır (exceljs "comments"/"Target" hatasıyla düşer)
      const relsPath = relsPathOf(s.path);
      const relsXml = await this.text(relsPath);
      let rels = relsXml ? parseRels(relsXml) : null;
      const exists = (r: Rel) => r.external || relKind(r.type) === "hyperlink" || !!this.zip.file(resolvePath(s.path, r.target));
      if (rels && relsXml && rels.some((r) => !exists(r))) {
        for (const r of rels) if (!exists(r) && LOST_LABEL[relKind(r.type)]) this.warn(LOST_LABEL[relKind(r.type)]);
        const keep = new Set(rels.filter(exists).map((r) => r.id));
        this.put(
          relsPath,
          relsXml.replace(/<(?:\w+:)?Relationship\b[^>]*>/g, (tag) => (keep.has(attrOf(tag, "Id") ?? "") ? tag : "")),
        );
        rels = rels.filter((r) => keep.has(r.id));
      }
      const refs = dropDanglingRefs(xml, rels, exists);
      if (refs) {
        xml = refs.xml;
        for (const l of refs.lost) this.warn(l);
      }
      // Birleştirmeler en son: exceljs kopyası sayfanın son hâlinden çıkarılır
      const merges = repairMerges(xml);
      if (merges) {
        xml = merges.xml;
        if (merges.merges.length && !this.merges.has(s.name)) this.merges.set(s.name, merges.merges);
        if (merges.dropped) this.warn("geçersiz ya da çakışan birleştirmeler");
        if (merges.big) this.warn("çok büyük birleştirilmiş alanlar");
        this.excelSheets.set(s.path, merges.excel);
      }
      if (!this.hadCells && /<(?:\w+:)?c\b[^>]*>\s*<(?:\w+:)?(?:v|is|f)\b/.test(xml)) this.hadCells = true;
      if (!hasSst && xml.includes(' t="s"')) {
        for (const m of xml.matchAll(/<(?:\w+:)?c\b[^>]*\st="s"[^>]*>\s*<(?:\w+:)?v>(\d+)</g)) maxSst = Math.max(maxSst, Math.min(Number(m[1]), 1_000_000));
      }
      if (!hasStyles && xml.includes("dxfId=")) for (const m of xml.matchAll(/\sdxfId="(\d+)"/g)) maxDxf = Math.max(maxDxf, Math.min(Number(m[1]), 10_000));
      if (xml !== orig) this.put(s.path, xml);
    }
    // Çizimlerin kopuk grafik/resim hedefleri (ilişki dosyası da silinmiş olabilir)
    for (const name of Object.keys(this.zip.files)) {
      if (!/^xl\/drawings\/[^/]+\.xml$/.test(name)) continue;
      const relsXml = await this.text(relsPathOf(name));
      if (relsXml === undefined) {
        const dx = (await this.text(name)) ?? "";
        if (/\s\w+:(?:id|embed)="/.test(dx)) this.warn(LOST_LABEL.drawing);
        continue;
      }
      for (const r of parseRels(relsXml)) {
        const kind = relKind(r.type);
        if (!r.external && LOST_LABEL[kind] && !this.zip.file(resolvePath(name, r.target))) this.warn(LOST_LABEL[kind]);
      }
    }
    if (!hasStyles) {
      // Stil bölümü yok: en küçük stil sayfası (koşullu biçimler için yeterli sayıda boş dxf ile;
      // exceljs dxf listesi yokken "getDxfStyle" hatasıyla dosyanın tamamını reddeder)
      this.put("xl/styles.xml", minStyles(maxDxf + 1), "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
      this.warn("hücre biçimleri (stil bölümü eksikti)");
    }
    if (maxSst >= 0) {
      // Paylaşılan metin bölümü yok: metin hücreleri sayı (dizin) olarak görünmesin ya da exceljs düşmesin, boş kalır
      this.put("xl/sharedStrings.xml", emptySst(maxSst + 1), "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
      this.warn("metin hücrelerinin içeriği (paylaşılan metin bölümü eksikti)");
      this.notices.push("Dosya bozuk: metin hücrelerinin içeriği dosyada yok; bu hücreler boş açıldı. Kaydetmeden önce içeriği denetleyin.");
    }
    // Paylaşılan metinlerde DOCTYPE/varlık tanımı (exceljs "undefined entity" ile düşer)
    if (hasSst) {
      const sst = await this.text("xl/sharedStrings.xml");
      const fixed = sst !== undefined ? stripDoctype(sst) : null;
      if (fixed !== null) {
        this.put("xl/sharedStrings.xml", fixed);
        this.warn("tanımsız XML varlıkları (dosya değiştirilmiş olabilir)");
      }
    }
    // Aşırı sıkıştırılmış XML dışı parçalar (gömülü dosya bombası) atılır
    for (const path of this.suspicious) {
      if (/\.(xml|rels|vml)$/i.test(path) || !this.zip.file(path)) continue;
      this.zip.remove(path);
      this.changed = true;
      this.warn("olağan dışı büyük bir gömülü parça");
    }
  }

  /** Yeni parçaların içerik türleri. */
  async flushContentTypes(): Promise<void> {
    if (!this.pendingCt.length) return;
    let ct = await this.text("[Content_Types].xml");
    if (!ct) return;
    for (const [path, type] of this.pendingCt) {
      if (ct.includes(`PartName="/${path}"`)) continue;
      ct = ct.replace(/<\/(?:\w+:)?Types>\s*$/, (end) => `<Override PartName="/${path}" ContentType="${type}"/>${end}`);
    }
    this.zip.file("[Content_Types].xml", ct);
  }
}

const generate = (zip: JSZip) => zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE", compressionOptions: { level: 1 } });

/** XLSX paketini açar, denetler ve onarır. Ölümcül sorunlarda anlaşılır FormatError fırlatır. */
export async function prepareXlsx(input: ArrayBuffer): Promise<PreparedXlsx> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(input);
  } catch (e) {
    throw toOpenError(e);
  }
  const suspicious = checkPackageLimits(zip);
  let changed = false;
  // Aşırı sıkıştırılmış (bomba şüphesi) XML parçaları her şeyden önce ayıklanır
  for (const path of suspicious) {
    if (!/\.(xml|rels|vml)$/i.test(path)) continue;
    try {
      zip.file(path, await readCollapsed(zip.file(path)!), { compression: "STORE" });
      changed = true;
    } catch (e) {
      throw toOpenError(e);
    }
  }
  memoizeText(zip);
  try {
    changed = (await normalizeZip(zip)) || changed;
  } catch {
    /* normalleştirme isteğe bağlı */
  }
  memoizeText(zip);
  const g = new Guard(zip, suspicious);
  try {
    await g.repairWorkbook();
    await g.repairSheets();
    await g.flushContentTypes();
  } catch (e) {
    throw toOpenError(e);
  }
  changed = changed || g.changed;
  let data = input;
  if (changed) {
    try {
      data = await generate(zip);
    } catch (e) {
      throw toOpenError(e);
    }
  }
  // exceljs kopyası: tablo parçaları zararsız taslakla değiştirilir, birleştirmeler çıkarılır
  let excel = data;
  const tables = Object.keys(zip.files).filter((n) => /^xl\/tables\/[^/]+\.xml$/i.test(n));
  if (tables.length || g.excelSheets.size) {
    const copy = await JSZip.loadAsync(data);
    tables.forEach((p, i) => copy.file(p, TABLE_STUB(i + 1)));
    for (const [p, x] of g.excelSheets) copy.file(p, x, { compression: x.length > 4 * MB ? "STORE" : "DEFLATE" });
    excel = await generate(copy);
  }
  return { data, excel, zip, merges: g.merges, warnings: g.warnings, notice: g.notices.length ? g.notices.join(" ") : undefined, hadCells: g.hadCells, sheetNames: g.sheets.map((s) => s.name) };
}
