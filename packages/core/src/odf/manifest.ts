// ODF paketi: mimetype ailesi ve META-INF/manifest.xml (OASIS ODF 1.3 Bölüm 2 §4). Okuyucular açılan paketin
// manifest listesini buradan alır; yazıcılar yeniden üretilen manifeste taşınan dosyaları ekler.
import type { OpcPackage } from "../opc/package";

export const MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
export const ODF_MIME_PREFIX = "application/vnd.oasis.opendocument.";

/** Belge ailesi (şablonlar dahil: "…text-template" → "text"). */
export type OdfFamily = "text" | "spreadsheet" | "presentation" | "graphics" | "chart" | "formula" | "image" | "database" | "other";

const FAMILIES: readonly OdfFamily[] = ["text", "spreadsheet", "presentation", "graphics", "chart", "formula", "image", "database"];

/** ODF içerik türünün ailesi; ODF değilse null. */
export function odfFamily(mime: string | null | undefined): OdfFamily | null {
  const m = (mime ?? "").trim();
  if (!m.startsWith(ODF_MIME_PREFIX)) return null;
  const rest = m.slice(ODF_MIME_PREFIX.length).replace(/-template$/, "");
  return (FAMILIES as readonly string[]).includes(rest) ? (rest as OdfFamily) : "other";
}

export interface ManifestEntry {
  /** Paket içi yol ("content.xml", alt belge için "Object 1/", kök için "/"). */
  path: string;
  mediaType: string;
  version?: string;
}

const ENT: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
const unesc = (s: string) =>
  s.replace(/&(#x[0-9a-fA-F]+|#\d+|\w+);/g, (m, e: string) => {
    if (e[0] !== "#") return ENT[e] ?? m;
    const n = e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
    return Number.isFinite(n) && n > 0 && n <= 0x10ffff ? String.fromCodePoint(n) : m;
  });
export const escAttr = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

/** Öznitelik değeri (ön ekten bağımsız yerel adla). */
function attrOf(tag: string, local: string): string | undefined {
  const m = new RegExp(`\\s(?:[\\w.-]+:)?${local}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(tag);
  return m ? unesc(m[1] ?? m[2] ?? "") : undefined;
}

/** manifest.xml → girdiler (şifreleme bilgisi yok sayılır; bozuk girdiler atlanır). */
export function parseManifest(xml: string | null | undefined): ManifestEntry[] {
  const out: ManifestEntry[] = [];
  for (const m of (xml ?? "").matchAll(/<(?:[\w.-]+:)?file-entry\b[^>]*>/g)) {
    const path = attrOf(m[0], "full-path");
    if (path === undefined || path === "") continue;
    const version = attrOf(m[0], "version");
    out.push({ path, mediaType: attrOf(m[0], "media-type") ?? "", ...(version ? { version } : {}) });
  }
  return out;
}

/** Girdilerden manifest.xml (kök girdisi "/" listede yoksa başa eklenmez; çağıran verir). */
export function manifestXml(entries: readonly ManifestEntry[], version = "1.3"): string {
  const rows = entries.map((e) => `<manifest:file-entry manifest:full-path="${escAttr(e.path)}"${e.version ? ` manifest:version="${escAttr(e.version)}"` : ""} manifest:media-type="${escAttr(e.mediaType)}"/>`);
  return `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="${MANIFEST_NS}" manifest:version="${version}">${rows.join("")}</manifest:manifest>`;
}

export interface OdfInfo {
  /** mimetype girdisi (yoksa manifest kök girdisinin türü). */
  mime: string;
  family: OdfFamily;
  manifest: ManifestEntry[];
  /** Manifestte şifreleme bilgisi var. */
  encrypted: boolean;
}

const infoCache = new WeakMap<OpcPackage, Promise<OdfInfo | null>>();

/**
 * Paketin ODF bilgisi (önbellekli). OPC paketi ya da ODF olmayan ZIP için null; bozuk manifest hata değildir
 * (boş liste).
 */
export function odfInfo(pkg: OpcPackage | null | undefined): Promise<OdfInfo | null> {
  if (!pkg) return Promise.resolve(null);
  let p = infoCache.get(pkg);
  if (!p) {
    p = (async (): Promise<OdfInfo | null> => {
      if (pkg.isOpc) return null;
      const read = (name: string) => pkg.text(name).catch(() => null);
      const manifestText = await read("META-INF/manifest.xml");
      const manifest = parseManifest(manifestText);
      const mime = ((await read("mimetype")) ?? manifest.find((e) => e.path === "/")?.mediaType ?? "").trim();
      const family = odfFamily(mime);
      if (!family) return null;
      return { mime, family, manifest, encrypted: /encryption-data/.test(manifestText ?? "") };
    })();
    infoCache.set(pkg, p);
  }
  return p;
}

/** Paketin ODF olup olmadığına hızlı (eşzamanlı) bakış: OPC değil ve içerik ya da mimetype girdisi var. */
export function looksOdf(pkg: OpcPackage | null | undefined): pkg is OpcPackage {
  return !!pkg && !pkg.isOpc && (pkg.has("content.xml") || pkg.has("mimetype"));
}

/** Uzantıdan tahmini içerik türü (kaynak manifestte girdisi olmayan dosyalar için). */
export function guessMediaType(path: string): string {
  const ext = /\.([\w]+)$/.exec(path)?.[1]?.toLowerCase() ?? ""; // yerel-ayar: ascii
  const map: Record<string, string> = { xml: "text/xml", rdf: "application/rdf+xml", png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", svg: "image/svg+xml", bmp: "image/bmp", tif: "image/tiff", tiff: "image/tiff", wmf: "image/wmf", emf: "image/emf", webp: "image/webp" };
  return map[ext] ?? "";
}
