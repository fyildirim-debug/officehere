// OPC paketi (ECMA-376 Bölüm 2): içerik türleri, ilişkiler ve parça erişimi.
// Okuma tembeldir: açılışta yalnız merkez dizin, [Content_Types].xml, kök ve ana parça ilişkileri
// açılır; parçalar istendiğinde açılır. Bozuk .rels ya da içerik türleri hata değildir (boş sayılır).

import { inflateEntry, readZip, toBytes, type ZipEntry } from "./zip";
import { isMacroContentType, relTypeName } from "./policy";

export const CONTENT_TYPES_PART = "[Content_Types].xml";
export const NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
export const NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";
/** XML'de ilişki kimliği taşıyan özniteliklerin ad alanları (Transitional ve Strict). */
export const NS_R = ["http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships"] as const;
const NS_VML_OFFICE = "urn:schemas-microsoft-com:office:office";

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** Varsayılan sınırlar: parça açık boyutu, .rels ve içerik türleri metni. */
export const OPC_LIMITS = { maxPartSize: 256 * 1024 * 1024, maxRelsSize: 16 * 1024 * 1024 };

/* ------------------------------------------------------------ küçük yardımcılar */

export function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

export function unescapeXml(s: string): string {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|quot|apos|lt|gt|amp);/g, (m, e: string) => {
    if (e[0] === "#") {
      const n = e[1] === "x" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return n > 0 && n <= 0x10ffff ? String.fromCodePoint(n) : m;
    }
    return { quot: '"', apos: "'", lt: "<", gt: ">", amp: "&" }[e] ?? m;
  });
}

/** Açılış etiketinin öznitelikleri (ad → çözülmüş değer). */
export function parseAttrs(s: string): Map<string, string> {
  const out = new Map<string, string>();
  for (const m of s.matchAll(/([\w.:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) if (!out.has(m[1])) out.set(m[1], unescapeXml(m[2] ?? m[3] ?? ""));
  return out;
}

function percentDecode(s: string): string {
  if (!s.includes("%")) return s;
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}

const SAFE_SEGMENT = /^[A-Za-z0-9._~!$&'()*+,;=:@\-[\]]*$/;

/** Parça adını ilişki hedefi olarak yazılabilir URI'ye çevirir (yalnız gereken bölümler kodlanır). */
export function encodePartUri(name: string): string {
  return name
    .split("/")
    .map((s) => (SAFE_SEGMENT.test(s) ? s : encodeURIComponent(s)))
    .join("/");
}

/** Metin kodlaması: UTF-8 (BOM'lu/BOM'suz) ya da UTF-16 BOM'lu. */
export function decodeText(b: Uint8Array): string {
  if (b.length >= 2 && b[0] === 0xff && b[1] === 0xfe) return new TextDecoder("utf-16le").decode(b.subarray(2));
  if (b.length >= 2 && b[0] === 0xfe && b[1] === 0xff) return new TextDecoder("utf-16be").decode(b.subarray(2));
  if (b.length >= 3 && b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return new TextDecoder("utf-8").decode(b.subarray(3));
  return new TextDecoder("utf-8").decode(b);
}

/* ------------------------------------------------------------ yollar */

/** Parçanın ilişki dosyası ("" = paket kökü → "_rels/.rels"). */
export function relsPathOf(part: string): string {
  if (part === "") return "_rels/.rels";
  const i = part.lastIndexOf("/");
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
}

/** İlişki dosyasının kaynak parçası ("_rels/.rels" → ""); ilişki dosyası değilse null. */
export function partOfRels(relsPath: string): string | null {
  if (relsPath.toLowerCase() === "_rels/.rels") return ""; // yerel-ayar: ascii
  const m = /^(.*?)_rels\/([^/]+)\.rels$/i.exec(relsPath);
  return m ? m[1] + m[2] : null;
}

/** Paket parçası olamayacak adlar: içerik türleri ve ilişki dosyaları. */
export function isReservedName(name: string): boolean {
  const k = name.toLowerCase(); // yerel-ayar: ascii
  return k === CONTENT_TYPES_PART.toLowerCase() || k.endsWith(".rels") || /(^|\/)_rels\//.test(k); // yerel-ayar: ascii
}

/**
 * İlişki hedefini kaynak parçanın dizinine göre paket içi ada çevirir. Parça bölümü (#…) atılır,
 * yüzde kodlaması çözülür. Kökün dışına çıkan, boş ya da şema içeren (http:) hedefte null.
 */
export function resolveTarget(source: string, target: string): string | null {
  const t = percentDecode(target.split("#")[0]).replace(/\\/g, "/");
  if (!t || /^[a-z][a-z0-9+.-]*:/i.test(t)) return null;
  const segs = t.startsWith("/") ? [] : source.split("/").slice(0, -1);
  for (const s of t.split("/")) {
    if (s === "..") {
      if (!segs.length) return null;
      segs.pop();
    } else if (s !== "." && s !== "") segs.push(s);
  }
  return segs.length ? segs.join("/") : null;
}

/** Kaynak parçadan hedef parçaya göreli ilişki hedefi (kök kaynakta paket içi ad). */
export function relativeTarget(source: string, part: string): string {
  const from = source.split("/").slice(0, -1);
  const to = part.split("/");
  let i = 0;
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
  return encodePartUri([...from.slice(i).map(() => ".."), ...to.slice(i)].join("/"));
}

/* ------------------------------------------------------------ içerik türleri */

/** [Content_Types].xml modeli. Parça adları baştaki "/" olmadan verilir; eşleme büyük/küçük harf duyarsızdır. */
export class ContentTypes {
  private defaults = new Map<string, string>();
  private overrides = new Map<string, { name: string; type: string }>();

  static parse(xml: string | null | undefined): ContentTypes {
    const ct = new ContentTypes();
    if (!xml) return ct;
    for (const m of xml.matchAll(/<(?:[\w.-]+:)?Default\b([^>]*)>/g)) {
      const a = parseAttrs(m[1]);
      const e = a.get("Extension");
      const t = a.get("ContentType");
      if (e && t && !ct.defaults.has(e.toLowerCase())) ct.defaults.set(e.toLowerCase(), t); // yerel-ayar: ascii
    }
    for (const m of xml.matchAll(/<(?:[\w.-]+:)?Override\b([^>]*)>/g)) {
      const a = parseAttrs(m[1]);
      const p = a.get("PartName");
      const t = a.get("ContentType");
      if (!p || !t) continue;
      const name = percentDecode(p.replace(/^\/+/, ""));
      if (!ct.overrides.has(name.toLowerCase())) ct.overrides.set(name.toLowerCase(), { name, type: t }); // yerel-ayar: ascii
    }
    return ct;
  }

  clone(): ContentTypes {
    const c = new ContentTypes();
    c.defaults = new Map(this.defaults);
    c.overrides = new Map(this.overrides);
    return c;
  }

  /** Parçanın içerik türü: önce Override, sonra uzantı varsayılanı. */
  get(part: string): string | undefined {
    return this.overrides.get(part.toLowerCase())?.type ?? this.defaults.get(extOf(part)); // yerel-ayar: ascii
  }

  getOverride(part: string): string | undefined {
    return this.overrides.get(part.toLowerCase())?.type; // yerel-ayar: ascii
  }

  getDefault(ext: string): string | undefined {
    return this.defaults.get(ext.toLowerCase()); // yerel-ayar: ascii
  }

  setOverride(part: string, type: string): void {
    this.overrides.set(part.toLowerCase(), { name: part, type }); // yerel-ayar: ascii
  }

  deleteOverride(part: string): void {
    this.overrides.delete(part.toLowerCase()); // yerel-ayar: ascii
  }

  setDefault(ext: string, type: string): void {
    this.defaults.set(ext.toLowerCase(), type); // yerel-ayar: ascii
  }

  /** Override girdileri: [parça adı, tür]. */
  overrideEntries(): [string, string][] {
    return [...this.overrides.values()].map((o) => [o.name, o.type]);
  }

  defaultEntries(): [string, string][] {
    return [...this.defaults];
  }

  toXml(): string {
    const d = [...this.defaults].map(([e, t]) => `<Default Extension="${escapeXml(e)}" ContentType="${escapeXml(t)}"/>`).join("");
    const o = [...this.overrides.values()].map((x) => `<Override PartName="/${escapeXml(encodePartUri(x.name))}" ContentType="${escapeXml(x.type)}"/>`).join("");
    return `${XML_DECL}<Types xmlns="${NS_CONTENT_TYPES}">${d}${o}</Types>`;
  }
}

function extOf(part: string): string {
  const slash = part.lastIndexOf("/");
  const dot = part.lastIndexOf(".");
  return dot > slash ? part.slice(dot + 1).toLowerCase() : ""; // yerel-ayar: ascii
}

/* ------------------------------------------------------------ ilişkiler */

export interface Relationship {
  id: string;
  /** Tam tür adresi. */
  type: string;
  /** Dosyada yazıldığı gibi hedef. */
  target: string;
  external: boolean;
  /** İç hedefin çözülmüş parça adı; dış hedefte ya da çözülemeyende null. */
  part: string | null;
}

/** .rels metnini ayrıştırır. Bozuk metinde okunabilen ilişkiler döner; Id/Type/Target eksik olan ve yinelenen Id atlanır. */
export function parseRels(xml: string | null | undefined, source: string): Relationship[] {
  const out: Relationship[] = [];
  if (!xml) return out;
  const ids = new Set<string>();
  for (const m of xml.matchAll(/<(?:[\w.-]+:)?Relationship\b([^>]*)>/g)) {
    const a = parseAttrs(m[1]);
    const id = a.get("Id");
    const type = a.get("Type");
    const target = a.get("Target");
    if (!id || !type || target === undefined || ids.has(id)) continue;
    ids.add(id);
    const external = (a.get("TargetMode") ?? "").toLowerCase() === "external"; // yerel-ayar: ascii
    out.push({ id, type, target, external, part: external ? null : resolveTarget(source, target) });
  }
  return out;
}

/** İlişki listesinin .rels metni. */
export function relsXml(rels: readonly Pick<Relationship, "id" | "type" | "target" | "external">[]): string {
  const body = rels.map((r) => `<Relationship Id="${escapeXml(r.id)}" Type="${escapeXml(r.type)}" Target="${escapeXml(r.target)}"${r.external ? ' TargetMode="External"' : ""}/>`).join("");
  return `${XML_DECL}<Relationships xmlns="${NS_RELATIONSHIPS}">${body}</Relationships>`;
}

/**
 * Bir parçanın ilişki listesi kurucusu. Yeni kimlikler var olanlardan (ve daha önce verilenlerden)
 * hiçbiriyle çakışmaz; eski bir kimlik asla yeniden kullanılmaz.
 */
export class RelsBuilder {
  readonly source: string;
  private list: Relationship[];
  private used: Set<string>;
  private next = 1;

  constructor(source: string, existing: readonly Relationship[] = []) {
    this.source = source;
    this.list = existing.map((r) => ({ ...r }));
    this.used = new Set(this.list.map((r) => r.id));
  }

  /** İç parçaya ilişki ekler (hedef paket içi mutlak adla verilir); yeni kimliği döner. */
  addPart(type: string, part: string): string {
    return this.push({ type, target: relativeTarget(this.source, part), external: false, part });
  }

  /** Dış adrese ilişki ekler. */
  addExternal(type: string, url: string): string {
    return this.push({ type, target: url, external: true, part: null });
  }

  private push(r: Omit<Relationship, "id">): string {
    let id: string;
    do id = `rId${this.next++}`;
    while (this.used.has(id));
    this.used.add(id);
    this.list.push({ id, ...r });
    return id;
  }

  get relationships(): readonly Relationship[] {
    return this.list;
  }

  get size(): number {
    return this.list.length;
  }

  xml(): string {
    return relsXml(this.list);
  }
}

const reEsc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** XML'de ilişki ad alanına (ve VML o:relid'e) bağlı önekler; "r" ve "o" her zaman dahil (bildirimsiz parçalar için). */
function relPrefixes(xml: string, extra: readonly string[] = []): { r: string[]; o: string[] } {
  const r = new Set(["r", ...extra]);
  const o = new Set(["o"]);
  for (const m of xml.matchAll(/\bxmlns:([\w.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) {
    const uri = m[2] ?? m[3] ?? "";
    if ((NS_R as readonly string[]).includes(uri)) r.add(m[1]);
    else if (uri === NS_VML_OFFICE) o.add(m[1]);
  }
  return { r: [...r], o: [...o] };
}

function relAttrRegex(xml: string, extra?: readonly string[]): RegExp {
  const { r, o } = relPrefixes(xml, extra);
  // r:id, r:embed, r:link, r:dm, r:pict…  |  o:relid (VML)  |  relId (dsp:dataModelExt, öneksiz)
  const alt = [...r.map((p) => `${reEsc(p)}:[A-Za-z_][\\w.-]*`), ...o.map((p) => `${reEsc(p)}:relid`), "relId"].join("|");
  return new RegExp(`(\\s(?:${alt})\\s*=\\s*)(?:"([^"]*)"|'([^']*)')`, "g");
}

/**
 * XML'de geçen ilişki kimlikleri (açık ilişkiler). İlişki ad alanına bağlı öneklerin bütün öznitelikleri
 * (r:id, r:embed, r:dm…), VML o:relid ve öneksiz relId sayılır. `prefixes`: bildirimi parçanın dışında
 * kalmış XML parçacıkları için ek önekler.
 */
export function explicitRelIds(xml: string, opts: { prefixes?: readonly string[] } = {}): Set<string> {
  const out = new Set<string>();
  for (const m of xml.matchAll(relAttrRegex(xml, opts.prefixes))) {
    const v = unescapeXml(m[2] ?? m[3] ?? "").trim();
    if (v) out.add(v);
  }
  return out;
}

/** XML'deki ilişki kimliklerini eşlemeye göre yeniden yazar (tek geçişte; rId1↔rId2 takası güvenli). */
export function rewriteRelIds(xml: string, map: ReadonlyMap<string, string>, opts: { prefixes?: readonly string[] } = {}): string {
  if (!map.size) return xml;
  return xml.replace(relAttrRegex(xml, opts.prefixes), (m, head: string, dq?: string, sq?: string) => {
    const v = unescapeXml(dq ?? sq ?? "").trim();
    const n = map.get(v);
    return n === undefined ? m : `${head}"${escapeXml(n)}"`;
  });
}

/* ------------------------------------------------------------ paket */

export interface OpcOpenOptions {
  /** Açılacak parçanın en büyük açık boyutu (bayt). Varsayılan 256 MB. */
  maxPartSize?: number;
}

/** Paketin ilişki çizgesi: parça → ilişkileri (tembel, önbellekli), erişilebilirlik. */
export class RelsGraph {
  private cache = new Map<string, Promise<Relationship[]>>();

  constructor(private readonly pkg: OpcPackage) {}

  /** Kaynak parçanın ilişkileri ("" = paket kökü). .rels yoksa ya da bozuksa boş liste. */
  rels(source: string): Promise<Relationship[]> {
    const key = source.toLowerCase(); // yerel-ayar: ascii
    let p = this.cache.get(key);
    if (!p) {
      p = this.pkg.textLimited(relsPathOf(source), OPC_LIMITS.maxRelsSize).then(
        (x) => parseRels(x, source),
        () => [],
      );
      this.cache.set(key, p);
    }
    return p;
  }

  /** Kaynaktan iç hedefi pakette bulunan ilişkiler. */
  async targets(source: string): Promise<Relationship[]> {
    return (await this.rels(source)).filter((r) => r.part !== null && this.pkg.has(r.part));
  }

  /**
   * Köklerden ilişkilerle ulaşılan parçalar (döngü korumalı). Kökler varsayılan olarak paket kökü.
   * Dönen kümede parça adları paketteki yazımıyla bulunur.
   */
  async reachable(roots: readonly string[] = [""]): Promise<Set<string>> {
    const seen = new Set<string>();
    const queue = [...roots];
    while (queue.length) {
      const batch = queue.splice(0, queue.length);
      const lists = await Promise.all(batch.map((s) => this.targets(s)));
      for (const list of lists) {
        for (const r of list) {
          const name = this.pkg.canonicalName(r.part!)!;
          if (seen.has(name)) continue;
          seen.add(name);
          queue.push(name);
        }
      }
    }
    return seen;
  }
}

/**
 * Açılmış OPC paketi. Baytlar bellekte tutulur, değiştirilmez; parçalar istendiğinde açılır.
 * Ham kopya için girdiler (`entry`) ve kaynak baytlar (`bytes`) doğrudan kullanılabilir.
 */
export class OpcPackage {
  readonly bytes: Uint8Array;
  readonly entries: readonly ZipEntry[];
  /** Okunamayan ya da yinelenen ZIP girdisi sayısı. */
  readonly skippedEntries: number;
  readonly graph: RelsGraph;
  readonly maxPartSize: number;
  private byName = new Map<string, ZipEntry>();
  private explicitCache = new Map<string, Promise<Set<string>>>();
  private ct = new ContentTypes();
  private main: string | null = null;
  private rootRels: Relationship[] = [];
  private mainRels: Relationship[] = [];

  private constructor(bytes: Uint8Array, entries: ZipEntry[], skipped: number, maxPartSize: number) {
    this.bytes = bytes;
    this.entries = entries;
    this.skippedEntries = skipped;
    for (const e of entries) this.byName.set(e.name.toLowerCase(), e); // yerel-ayar: ascii
    this.maxPartSize = maxPartSize;
    this.graph = new RelsGraph(this);
  }

  /**
   * Paketi açar: merkez dizin, içerik türleri, kök ve ana parça ilişkileri. ZIP değilse OpcError
   * ("not-zip" / "corrupt"); [Content_Types].xml eksik ya da bozuksa hata değildir (`isOpc` false).
   */
  static async open(data: Uint8Array | ArrayBuffer, opts: OpcOpenOptions = {}): Promise<OpcPackage> {
    const zip = readZip(toBytes(data));
    const pkg = new OpcPackage(zip.bytes, zip.entries, zip.skipped, opts.maxPartSize ?? OPC_LIMITS.maxPartSize);
    const [ctXml, rootRels] = await Promise.all([pkg.textLimited(CONTENT_TYPES_PART, OPC_LIMITS.maxRelsSize).catch(() => null), pkg.graph.rels("")]);
    pkg.ct = ContentTypes.parse(ctXml);
    pkg.rootRels = rootRels;
    const main = rootRels.find((r) => relTypeName(r.type) === "officeDocument" && r.part && pkg.has(r.part));
    pkg.main = main ? pkg.canonicalName(main.part!)! : null;
    if (pkg.main) pkg.mainRels = await pkg.graph.rels(pkg.main);
    return pkg;
  }

  /** [Content_Types].xml modeli (değiştirmeden önce `clone()` alın). */
  get contentTypes(): ContentTypes {
    return this.ct;
  }

  /** Ana parça (kökteki officeDocument ilişkisinin hedefi); yoksa null. */
  get mainPart(): string | null {
    return this.main;
  }

  /** [Content_Types].xml var mı (OOXML paketi mi)? */
  get isOpc(): boolean {
    return this.byName.has(CONTENT_TYPES_PART.toLowerCase()); // yerel-ayar: ascii
  }

  /** Parça adları (arşiv sırasıyla, [Content_Types].xml ve .rels dahil). */
  get names(): string[] {
    return this.entries.map((e) => e.name);
  }

  has(name: string): boolean {
    return this.byName.has(name.toLowerCase()); // yerel-ayar: ascii
  }

  entry(name: string): ZipEntry | undefined {
    return this.byName.get(name.toLowerCase()); // yerel-ayar: ascii
  }

  /** Büyük/küçük harf duyarsız adın paketteki yazımı. */
  canonicalName(name: string): string | undefined {
    return this.byName.get(name.toLowerCase())?.name; // yerel-ayar: ascii
  }

  contentType(name: string): string | undefined {
    return this.contentTypes.get(this.canonicalName(name) ?? name);
  }

  /** Parçanın açık boyutu (merkez dizinden); yoksa -1. */
  sizeOf(name: string): number {
    return this.entry(name)?.size ?? -1;
  }

  /** Parçanın baytları; yoksa null. Çok büyük ya da bozuk parçada OpcError. Dönen dizi değiştirilmemeli. */
  async data(name: string): Promise<Uint8Array | null> {
    const e = this.entry(name);
    return e ? inflateEntry(this.bytes, e, this.maxPartSize) : null;
  }

  /** Parçanın metni (UTF-8/UTF-16); yoksa null. */
  async text(name: string): Promise<string | null> {
    return this.textLimited(name, this.maxPartSize);
  }

  /** @internal Sınırlı metin okuma (.rels ve içerik türleri için daha küçük sınır). */
  async textLimited(name: string, limit: number): Promise<string | null> {
    const e = this.entry(name);
    return e ? decodeText(await inflateEntry(this.bytes, e, Math.min(limit, this.maxPartSize))) : null;
  }

  /** Kaynak parçanın ilişkileri ("" = kök). */
  rels(source: string): Promise<Relationship[]> {
    return this.graph.rels(this.canonicalName(source) ?? source);
  }

  /** Parça XML'inde geçen ilişki kimlikleri (önbellekli). Parça yoksa boş küme; okunamazsa OpcError. */
  explicitRelIds(part: string): Promise<Set<string>> {
    const key = part.toLowerCase(); // yerel-ayar: ascii
    let p = this.explicitCache.get(key);
    if (!p) {
      p = this.text(part).then((x) => (x ? explicitRelIds(x) : new Set<string>()));
      this.explicitCache.set(key, p);
    }
    return p;
  }

  get mainContentType(): string | undefined {
    return this.mainPart ? this.contentType(this.mainPart) : undefined;
  }

  /** Makrolu paket mi (ana parça türü makrolu ya da ana parçada vbaProject ilişkisi var)? */
  get macroEnabled(): boolean {
    return isMacroContentType(this.mainContentType) || this.mainRels.some((r) => relTypeName(r.type) === "vbaProject");
  }

  /** Dijital imzalı mı (kökte imza kaynağı ilişkisi ya da _xmlsignatures/ altında parça)? Kaydedince imza kalkar. */
  get signed(): boolean {
    return this.rootRels.some((r) => relTypeName(r.type) === "origin") || this.entries.some((e) => /^_xmlsignatures\//i.test(e.name));
  }
}
