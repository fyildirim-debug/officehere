/** OpenDocument sunusu (ODF 1.3) ortak yardımcıları: ad alanları, birimler, renkler. */
import type { Color } from "../../model/types";

export const ODF_NS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  draw: "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
  xlink: "http://www.w3.org/1999/xlink",
  dc: "http://purl.org/dc/elements/1.1/",
  meta: "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
  presentation: "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
  svg: "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
  smil: "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0",
  anim: "urn:oasis:names:tc:opendocument:xmlns:animation:1.0",
  manifest: "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
};

/** Kök öğelerde kullanılan ad alanı bildirimleri. */
export const ODF_NS_DECL = Object.entries(ODF_NS)
  .filter(([k]) => k !== "manifest")
  .map(([k, v]) => `xmlns:${k}="${v}"`)
  .join(" ");

export const ODP_MIME = "application/vnd.oasis.opendocument.presentation";
export const OTP_MIME = "application/vnd.oasis.opendocument.presentation-template";

export class OdpError extends Error {}

/** ODF uzunluğu ("2cm", "12pt", "1in"…) → 96 dpi piksel. Geçersizse undefined. */
export function lengthPx(v: string | undefined | null): number | undefined {
  if (!v) return undefined;
  const m = /^\s*([-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?)\s*(cm|mm|in|inch|pt|pc|px)?\s*$/i.exec(v);
  if (!m) return undefined;
  const n = Number(m[1]);
  if (!Number.isFinite(n) || Math.abs(n) > 1e7) return undefined;
  switch ((m[2] ?? "cm").toLowerCase()) {
    case "mm":
      return (n * 96) / 25.4;
    case "in":
    case "inch":
      return n * 96;
    case "pt":
      return (n * 96) / 72;
    case "pc":
      return n * 16;
    case "px":
      return n;
    default:
      return (n * 96) / 2.54;
  }
}

/** ODF uzunluğu → punto. */
export function lengthPt(v: string | undefined | null): number | undefined {
  const px = lengthPx(v);
  return px === undefined ? undefined : (px * 72) / 96;
}

const r4 = (n: number) => Math.round(n * 10000) / 10000;

/** Piksel → "1.2345cm". */
export const cm = (px: number) => `${r4((px * 2.54) / 96)}cm`;
/** Punto → "12pt". */
export const ptStr = (pt: number) => `${r4(pt)}pt`;

/** "50%" → 0.5. */
export function percent(v: string | undefined | null): number | undefined {
  if (!v) return undefined;
  const m = /^\s*([-+]?\d+(?:\.\d+)?)\s*%\s*$/.exec(v);
  return m ? Number(m[1]) / 100 : undefined;
}

/** "#rrggbb" → model rengi. */
export function odfColor(v: string | undefined | null): Color | undefined {
  if (!v) return undefined;
  const m = /^#([0-9a-f]{6})$/i.exec(v.trim());
  if (m) return { rgb: `#${m[1].toUpperCase()}` };
  const s = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(v.trim());
  return s ? { rgb: `#${s[1]}${s[1]}${s[2]}${s[2]}${s[3]}${s[3]}`.toUpperCase() } : undefined;
}

/** Açı ("30deg", "0.5rad", "300" = onda bir derece) → derece. */
export function angleDeg(v: string | undefined | null): number | undefined {
  if (!v) return undefined;
  const m = /^\s*([-+]?\d+(?:\.\d+)?)\s*(deg|rad|grad)?\s*$/i.exec(v);
  if (!m) return undefined;
  const n = Number(m[1]);
  const u = (m[2] ?? "").toLowerCase();
  if (u === "deg") return n;
  if (u === "rad") return (n * 180) / Math.PI;
  if (u === "grad") return n * 0.9;
  return n / 10;
}

/** Yazı tipi ailesi listesinden ilk ad (tırnaklar atılır). */
export function firstFamily(v: string | undefined | null): string | undefined {
  if (!v) return undefined;
  const first = v.split(",")[0]?.trim().replace(/^['"]|['"]$/g, "");
  return first || undefined;
}

/** ODF öznitelik dizgesi; tanımsızlar atlanır (değer kaçışlı). */
export function odfAttrs(o: Record<string, string | number | undefined | null | false>): string {
  let s = "";
  for (const [k, v] of Object.entries(o)) {
    if (v === undefined || v === null || v === false) continue;
    s += ` ${k}="${String(v).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]!)}"`;
  }
  return s;
}

/** Bayt dizisini base64'e çevirir. */
export function toBase64(bytes: Uint8Array): string {
  let s = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) s += String.fromCharCode(...bytes.subarray(i, i + chunk));
  return btoa(s);
}

export const IMAGE_MIME: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  svg: "image/svg+xml",
  webp: "image/webp",
  tif: "image/tiff",
  tiff: "image/tiff",
};

/** Tarayıcının gösterebildiği resim türleri. */
export const DISPLAYABLE = new Set(["image/png", "image/jpeg", "image/gif", "image/bmp", "image/svg+xml", "image/webp"]);
