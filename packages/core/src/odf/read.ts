// ODF okurken: gömülü nesne çerçevesinin tanınması, türü (manifestten), tarayıcıda çizilebilen yedek görüntüleri
// ve DOM'dan okuyanlar için opak nesne kaydı (çerçeve XML'i + andığı otomatik stiller).
import type { OpcPackage } from "../opc/package";
import type { OdfObject } from "./carry";
import { odfFamily, type OdfInfo } from "./manifest";
import { elementXml, styleRefs } from "./xml";

/** Çerçevedeki gömülü nesnenin başvurusu (draw:object / draw:object-ole); yoksa null. */
export function objectHref(frameXml: string): { href: string; ole: boolean } | null {
  const m = /<(?:[\w.-]+:)?object(-ole)?\b[^>]*?\s[\w.-]+:href\s*=\s*"([^"]*)"/.exec(frameXml);
  return m ? { href: m[2].replace(/^(\.\/)+/, "").replace(/\/+$/, ""), ole: !!m[1] } : null;
}

const LABELS: Record<string, string> = {
  chart: "Grafik",
  formula: "Formül",
  spreadsheet: "Hesap tablosu",
  text: "Metin belgesi",
  presentation: "Sunu",
  graphics: "Çizim",
  image: "Resim",
};

/** Gömülü nesnenin kullanıcıya gösterilen türü ("Grafik", "Formül", "OLE nesnesi"…). */
export function objectLabel(info: OdfInfo | null | undefined, frameXml: string): string {
  const o = objectHref(frameXml);
  if (!o) return "Gömülü nesne";
  if (o.ole) return "OLE nesnesi";
  let href = o.href;
  try {
    if (/%[0-9a-fA-F]{2}/.test(href)) href = decodeURIComponent(href);
  } catch {
    // olduğu gibi
  }
  const fam = odfFamily(info?.manifest.find((e) => e.path === `${href}/`)?.mediaType);
  return (fam && LABELS[fam]) || "Gömülü nesne";
}

const IMG: [string, (b: Uint8Array) => boolean][] = [
  ["image/png", (b) => b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47],
  ["image/jpeg", (b) => b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff],
  ["image/gif", (b) => b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38],
  ["image/webp", (b) => b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50],
  ["image/svg+xml", (b) => /^\s*(<\?xml[^>]*>\s*)?(<!--[\s\S]*?-->\s*)*(<!DOCTYPE[^>]*>\s*)?<svg[\s>]/.test(new TextDecoder().decode(b.subarray(0, 512)))],
];

function base64(b: Uint8Array): string {
  let s = "";
  for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000));
  return btoa(s);
}

/**
 * Yedek görüntülerden (ObjectReplacements/) tarayıcıda çizilebilenler: yol → data: adresi. LibreOffice'in
 * StarView metafile (SVM) yedekleri çizilemez; onlar için çağıran yer tutucu gösterir.
 */
export async function objectPreviews(pkg: OpcPackage, opts: { maxBytes?: number; maxCount?: number } = {}): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  const max = opts.maxBytes ?? 8 * 1024 * 1024;
  let left = opts.maxCount ?? 64;
  for (const name of pkg.names) {
    if (!/^ObjectReplacements\//i.test(name) || pkg.sizeOf(name) > max || left-- <= 0) continue;
    const b = await pkg.data(name).catch(() => null);
    if (!b || b.length < 12) continue;
    const mime = IMG.find(([, ok]) => ok(b))?.[0];
    if (mime) out.set(name, `data:${mime};base64,${base64(b)}`);
  }
  return out;
}

/** Çerçevenin yedek görüntüsünün yolu (draw:image xlink:href); yoksa null. */
export function replacementHref(frameXml: string): string | null {
  const m = /<(?:[\w.-]+:)?image\b[^>]*?\s[\w.-]+:href\s*=\s*"([^"]*)"/.exec(frameXml);
  if (!m) return null;
  let p = m[1].replace(/^(\.\/)+/, "");
  try {
    if (/%[0-9a-fA-F]{2}/.test(p)) p = decodeURIComponent(p);
  } catch {
    // olduğu gibi
  }
  return p;
}

/**
 * DOM çerçevesinden opak nesne: XML'i (ad alanı bildirimleriyle) ve andığı otomatik stiller (`autoStyles`:
 * çerçevenin bulunduğu parçanın office:automatic-styles öğesi).
 */
export function odfObjectFromElement(el: Element, autoStyles: Element | null | undefined, src?: string): OdfObject {
  const xml = elementXml(el);
  const names = styleRefs(xml);
  const styles: string[] = [];
  if (autoStyles && names.size) {
    for (let c = autoStyles.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      const s = c as Element;
      if (s.localName !== "style") continue;
      const name = s.getAttributeNS("urn:oasis:names:tc:opendocument:xmlns:style:1.0", "name") ?? s.getAttribute("style:name");
      if (name && names.has(name)) styles.push(elementXml(s));
    }
  }
  return { ...(src ? { src } : {}), xml, styles };
}
