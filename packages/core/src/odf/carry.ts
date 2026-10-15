// ODF geçiş katmanı (yazarken). Açılan .odt/.ods/.odp'nin özgün paketi kaynak kaydındadır (OpcPackage ile okunur:
// ODF ZIP'i OPC değildir, ilişki grafiği yoktur; yalnız ZIP okuyucu, ham kopya ve kaynak kaydı ortaktır).
//
// Paket düzeyi (aynı aileye kaydederken): settings.xml (görünüm ve uyumluluk ayarları) ve manifest.rdf (RDF üst
// verisi) özgün dosyadan taşınır; yazıcı settings.xml yazdıysa onunkiler geçerlidir, kaynağın diğer ayar
// kümeleri eklenir. Thumbnails/ taşınmaz (eski önizleme yanıltır).
// Öğe düzeyi: modelin tanımadığı çizim nesneleri (draw:object ile gömülü grafik/formül alt belgeleri, OLE) özgün
// draw:frame XML'iyle taşınır: başvurduğu dosyalar (Object N/, ObjectReplacements/…) yeni adla kaynak paketten
// ham kopyalanır, andığı otomatik stiller yeniden adlandırılır.
// META-INF/manifest.xml yazıcınınkiyle taşınan dosyaların girdilerinden yeniden kurulur; mimetype ilk ve
// sıkıştırmasızdır. Yazıcının ve taşınan girdiler yeniden sıkıştırılmaz.
import type { OpcPackage } from "../opc/package";
import { packageSources, type SourceRegistry } from "../opc/source";
import { attrOf, childElements, rebuildElement } from "../opc/xmlPart";
import { inflateEntry, readZip, ZipWriter, type ZipEntry } from "../opc/zip";
import { escAttr, guessMediaType, looksOdf, manifestXml, odfInfo, parseManifest, type ManifestEntry, type OdfFamily } from "./manifest";
import { renameStyle, renameStyleRefs, rootNamespaces, styleIdentity, withNamespaces, XLINK_NS } from "./xml";

/** Modelin taşıdığı opak ODF çizim nesnesi (bayt taşımaz; dosyalar kaynak paketten kopyalanır). */
export interface OdfObject {
  /** Kaynak paket kaydı kimliği (verilmezse yazıcının varsayılan kaynağı). */
  src?: string;
  /** Özgün çizim öğesi (draw:frame), kullandığı ad alanlarının bildirimleri kökte. */
  xml: string;
  /** Andığı otomatik stiller (style:style; ad alanı bildirimleriyle). */
  styles: string[];
}

export type OdfDropReason = "source" | "missing" | "thumbnail";

export interface OdfCarryReport {
  /** Kaynak paketten kopyalanan dosyalar (çıktıdaki adlarıyla). */
  carried: string[];
  /** Yazılan opak nesne sayısı. */
  objects: number;
  /** Taşınamayan ya da bilerek düşürülenler. */
  dropped: { what: string; reason: OdfDropReason }[];
}

interface CopyPlan {
  pkg: OpcPackage;
  from: string;
  to: string;
  dir: boolean;
}

const lc = (s: string) => s.toLowerCase(); // yerel-ayar: ascii
/** Yazıcıların kökte ürettiği ya da ODF'nin ayırdığı adlar: taşınan nesne bunlarla adlandırılmaz. */
const RESERVED = new Set(["mimetype", "content.xml", "styles.xml", "meta.xml", "settings.xml", "manifest.rdf", "meta-inf", "pictures", "media", "thumbnails", "objectreplacements", "configurations2", "basic", "scripts", "dialogs"]);
const unescXml = (s: string) => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");

/** Paket içi göreli başvuru → paket yolu ("./Object 1" → "Object 1"); dış/geçersiz başvuruda null. */
function packagePath(href: string): string | null {
  if (!href || /^[a-z][a-z0-9+.-]*:/i.test(href) || href.startsWith("#") || href.startsWith("/")) return null;
  let p = href.replace(/^(\.\/)+/, "").replace(/\/+$/, "");
  try {
    if (/%[0-9a-fA-F]{2}/.test(p)) p = decodeURIComponent(p);
  } catch {
    // çözülemeyen kaçış olduğu gibi kalır
  }
  if (!p || p.split("/").some((s) => s === ".." || s === ".")) return null;
  return p;
}

export class OdfCarry {
  private plans: CopyPlan[] = [];
  /** Bu yazmada ayrılan adlar (küçük harf; dizinler sondaki "/" olmadan). */
  private used = new Set<string>();
  private styleNames = new Map<string, string>();
  private styleSeq = 0;
  private fileSeq = 0;
  private pkgIds = new Map<OpcPackage, number>();
  private registry: SourceRegistry;
  readonly report: OdfCarryReport = { carried: [], objects: 0, dropped: [] };

  constructor(private opts: { family: OdfFamily; source?: string | OpcPackage | null; registry?: SourceRegistry }) {
    this.registry = opts.registry ?? packageSources;
  }

  /** Paket ya da kayıt kimliği → ODF paketi (bulunamaz ya da ODF değilse undefined). */
  private pkgOf(src: string | OpcPackage | null | undefined): OpcPackage | undefined {
    const pkg = typeof src === "string" ? this.registry.get(src) : (src ?? undefined);
    return looksOdf(pkg) ? pkg : undefined;
  }

  /** Varsayılan kaynak (ODF ise). */
  get source(): OpcPackage | undefined {
    return this.pkgOf(this.opts.source);
  }

  /** Taşınacak opak nesne var mı (ya da varsayılan kaynak ODF mi): `finish` gerekli mi. */
  get active(): boolean {
    return this.plans.length > 0 || !!this.source;
  }

  private allocate(name: string, reserved: boolean): string {
    let n = name;
    if (this.used.has(lc(n)) || (reserved && RESERVED.has(lc(n)))) {
      let k = 1;
      do n = `Object ${k++}`;
      while (this.used.has(lc(n)));
    }
    this.used.add(lc(n));
    return n;
  }

  /**
   * Opak nesneyi bu yazmaya ekler: başvurduğu dosyalar yeni adlarla kopyalanacak, XML'deki başvurular ve stil adları
   * yeniden yazılır. `scope`: stillerin yazılacağı otomatik stil bölümü (aynı stil her bölüme bir kez). Kaynak ya da
   * başvurulan dosya yoksa null (çağıran yedek görünümü yazar).
   */
  object(o: OdfObject, scope = ""): { xml: string; styles: string[] } | null {
    const pkg = this.pkgOf(o.src ?? this.opts.source);
    if (!pkg) {
      this.report.dropped.push({ what: "nesne", reason: "source" });
      return null;
    }
    const decls = rootNamespaces(o.xml);
    const xp = Object.entries(decls).find(([, u]) => u === XLINK_NS)?.[0] || "xlink";
    const hrefRe = new RegExp(`(\\s${xp.replace(/[.-]/g, "\\$&")}:href\\s*=\\s*")([^"]*)"`, "g");
    const names = pkg.names;
    const lower = names.map(lc);
    // Başvurular: önce nesnelerin kendisi (yedek görüntü adı nesneyi izlesin), sonra diğerleri.
    const refs = [...o.xml.matchAll(hrefRe)].map((m) => packagePath(unescXml(m[2]))).filter((p): p is string => !!p);
    const ordered = [...refs.filter((p) => !p.includes("/")), ...refs.filter((p) => p.includes("/"))];
    const map = new Map<string, string>();
    const plans: CopyPlan[] = [];
    const renamedObjects = new Map<string, string>();
    for (const path of ordered) {
      if (map.has(path)) continue;
      const lp = lc(path);
      const dir = lower.some((n) => n.startsWith(lp + "/"));
      const canonical = dir ? path : pkg.canonicalName(path);
      if (!dir && !canonical) {
        this.report.dropped.push({ what: path, reason: "missing" });
        return null;
      }
      let to: string;
      const slash = path.lastIndexOf("/");
      if (slash < 0) {
        to = this.allocate(path, true);
        renamedObjects.set(path, to);
      } else {
        const folder = path.slice(0, slash);
        const base = path.slice(slash + 1);
        const follow = lc(folder) === "objectreplacements" ? renamedObjects.get(base) : undefined;
        to = follow && !this.used.has(lc(`${folder}/${follow}`)) ? `${folder}/${follow}` : `${folder}/oh${++this.fileSeq}-${base}`;
        this.used.add(lc(to));
      }
      map.set(path, to);
      plans.push({ pkg, from: dir ? path : canonical!, to, dir });
    }
    this.plans.push(...plans);
    let xml = o.xml.replace(hrefRe, (m, lead: string, raw: string) => {
      const p = packagePath(unescXml(raw));
      const to = p ? map.get(p) : undefined;
      return to ? `${lead}${raw.startsWith("./") || !to.includes("/") ? "./" : ""}${escAttr(to)}"` : m;
    });
    // Stiller: kaynak paket + bölüm başına bir kez, yazıcının adlarıyla çakışmayan adla.
    let pid = this.pkgIds.get(pkg);
    if (pid === undefined) this.pkgIds.set(pkg, (pid = this.pkgIds.size + 1));
    const rename = new Map<string, [string, string][]>();
    const emit: string[] = [];
    for (const sx of o.styles) {
      const id = styleIdentity(sx);
      if (!id) continue;
      const key = `${pid}|${scope}|${id.family}|${id.name}`;
      let nn = this.styleNames.get(key);
      if (!nn) {
        nn = `oh${++this.styleSeq}-${id.name}`;
        this.styleNames.set(key, nn);
        emit.push(renameStyle(sx, nn));
      }
      const list = rename.get(id.name) ?? [];
      list.push([id.family, nn]);
      rename.set(id.name, list);
    }
    xml = renameStyleRefs(xml, rename);
    this.report.objects++;
    return { xml, styles: emit };
  }

  /**
   * Yazıcının çıktısına taşınanları ekler ve manifesti yeniden kurar. Taşınacak bir şey yoksa çıktı aynen döner.
   * Çıktı ZIP değilse OpcError.
   */
  async finish(output: Uint8Array): Promise<Uint8Array> {
    const src = this.source;
    const info = src ? await odfInfo(src) : null;
    const same = !!info && info.family === this.opts.family;
    if (!same && !this.plans.length) return output;
    const arch = readZip(output);
    const outNames = new Set(arch.entries.map((e) => lc(e.name)));
    const byName = new Map(arch.entries.map((e) => [e.name, e]));
    const z = new ZipWriter();
    const mime = byName.get("mimetype");
    if (mime) z.addRaw("mimetype", arch.bytes, mime);
    for (const e of arch.entries) if (e.name !== "mimetype" && e.name !== "META-INF/manifest.xml") z.addRaw(e.name, arch.bytes, e);
    const text = async (e: ZipEntry | undefined) => (e ? new TextDecoder().decode(await inflateEntry(arch.bytes, e, 64 * 1024 * 1024)) : null);
    const entries: ManifestEntry[] = parseManifest(await text(byName.get("META-INF/manifest.xml")));
    const listed = new Set(entries.map((e) => e.path));
    const addEntry = (path: string, mediaType: string, version?: string) => {
      if (listed.has(path)) return;
      listed.add(path);
      entries.push({ path, mediaType, ...(version ? { version } : {}) });
    };
    const copy = (pkg: OpcPackage, from: string, to: string): boolean => {
      const e = pkg.entry(from);
      if (!e || outNames.has(lc(to))) return false;
      z.addRaw(to, pkg.bytes, e);
      outNames.add(lc(to));
      this.report.carried.push(to);
      return true;
    };

    if (same && src && info) {
      const man = new Map(info.manifest.map((e) => [e.path, e]));
      if (src.has("settings.xml")) {
        if (!byName.has("settings.xml")) {
          if (copy(src, src.canonicalName("settings.xml")!, "settings.xml")) addEntry("settings.xml", "text/xml");
        } else {
          const merged = mergeSettings((await text(byName.get("settings.xml"))) ?? "", (await src.text("settings.xml").catch(() => null)) ?? "");
          if (merged) z.add("settings.xml", merged);
        }
      }
      // RDF üst verisi (manifest.rdf ve andığı üst veri dosyaları)
      const rdf = new Set(info.manifest.filter((e) => /^application\/rdf\+xml\b/.test(e.mediaType) && !e.path.endsWith("/")).map((e) => e.path));
      if (src.has("manifest.rdf")) rdf.add(src.canonicalName("manifest.rdf")!);
      for (const p of rdf) {
        if (p.includes("/") && /^(Thumbnails|META-INF)\//i.test(p)) continue;
        if (src.has(p) && copy(src, src.canonicalName(p)!, p)) addEntry(p, man.get(p)?.mediaType || "application/rdf+xml");
      }
      if (src.names.some((n) => /^Thumbnails\//i.test(n))) this.report.dropped.push({ what: "Thumbnails/", reason: "thumbnail" });
    }

    for (const p of this.plans) {
      const pinfo = await odfInfo(p.pkg);
      const man = new Map((pinfo?.manifest ?? []).map((e) => [e.path, e]));
      if (!p.dir) {
        if (copy(p.pkg, p.from, p.to)) addEntry(p.to, man.get(p.from)?.mediaType ?? guessMediaType(p.from), man.get(p.from)?.version);
        continue;
      }
      const prefix = p.from + "/";
      for (const name of p.pkg.names) {
        if (!name.startsWith(prefix)) continue;
        const to = p.to + "/" + name.slice(prefix.length);
        if (copy(p.pkg, name, to)) addEntry(to, man.get(name)?.mediaType ?? guessMediaType(name), man.get(name)?.version);
      }
      // Alt belgenin (ve iç içe alt belgelerin) dizin girdileri: içerik türü ve sürüm
      for (const e of pinfo?.manifest ?? []) if (e.path.startsWith(prefix) && e.path.endsWith("/")) addEntry(p.to + "/" + e.path.slice(prefix.length), e.mediaType, e.version);
      const d = man.get(prefix);
      addEntry(p.to + "/", d?.mediaType ?? "", d?.version);
    }
    z.add("META-INF/manifest.xml", manifestXml(entries));
    return z.finish();
  }
}

/**
 * settings.xml birleştirme: yazıcının ayar kümeleri (config:config-item-set) geçerli, kaynağın yazıcıda olmayan
 * kümeleri (ör. ooo:configuration-settings: yazıcı ayarları, uyumluluk) eklenir. Eklenecek yoksa null.
 */
export function mergeSettings(out: string, src: string): string | null {
  try {
    const outRoot = childElements(out);
    const outSettings = outRoot.find((c) => c.name === "settings");
    const srcSettings = childElements(src).find((c) => c.name === "settings");
    if (!outSettings || !srcSettings) return null;
    const outSets = childElements(outSettings.xml);
    const have = new Set(outSets.map((s) => attrOf(s.xml, "config:name")));
    const extra = childElements(srcSettings.xml).filter((s) => s.name === "config-item-set" && !have.has(attrOf(s.xml, "config:name")));
    if (!extra.length) return null;
    const settings = rebuildElement(outSettings.xml, [...outSets.map((s) => s.xml), ...extra.map((s) => s.xml)]);
    const root = rebuildElement(out, outRoot.map((c) => (c === outSettings ? settings : c.xml)));
    return `<?xml version="1.0" encoding="UTF-8"?>${withNamespaces(root, rootNamespaces(src))}`;
  } catch {
    return null;
  }
}
