// OPC testleri için yardımcılar: bellekte paket kurma ve paket tutarlılık denetimi.
import { readdirSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { ZipWriter } from "../../src/opc/zip";
// Derlemin paket denetimi (pnpm derlem:opc) testlerde de aynı ölçüt olsun diye doğrudan kullanılır.
// @ts-expect-error: tipsiz .mjs
import { opcDenetle } from "../../../../scripts/derlem/opc.mjs";

export const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const T = (short: string) => `${R}/${short}`;
export const MST = (path: string) => `http://schemas.microsoft.com/office/${path}`;
export const CT_MAIN_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
export const CT_MAIN_DOCM = "application/vnd.ms-word.document.macroEnabled.main+xml";

export interface RelSpec {
  id: string;
  type: string;
  target: string;
  external?: boolean;
}

export function rels(list: RelSpec[]): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list
    .map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"${r.external ? ' TargetMode="External"' : ""}/>`)
    .join("")}</Relationships>`;
}

export function contentTypes(defaults: Record<string, string>, overrides: Record<string, string>): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${Object.entries(defaults)
    .map(([e, t]) => `<Default Extension="${e}" ContentType="${t}"/>`)
    .join("")}${Object.entries(overrides)
    .map(([p, t]) => `<Override PartName="/${p}" ContentType="${t}"/>`)
    .join("")}</Types>`;
}

export const BASE_DEFAULTS = { rels: "application/vnd.openxmlformats-package.relationships+xml", xml: "application/xml" };

/** Parçalardan ZIP kurar (verilen sırayla). */
export async function zipOf(parts: Record<string, string | Uint8Array>): Promise<Uint8Array> {
  const z = new ZipWriter();
  for (const [name, data] of Object.entries(parts)) z.add(name, data);
  return z.finish();
}

export const DOC_XML = (body = "") =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}"><w:body>${body}<w:p/></w:body></w:document>`;

/** Yazıcı çıktısına benzer en küçük docx: document, styles, core; [Content_Types].xml bilerek sonda. */
export async function minimalDocx(extra: Record<string, string | Uint8Array> = {}, opts: { docRels?: RelSpec[]; overrides?: Record<string, string>; rootRels?: RelSpec[] } = {}): Promise<Uint8Array> {
  return zipOf({
    "_rels/.rels": rels([{ id: "rId1", type: T("officeDocument"), target: "word/document.xml" }, ...(opts.rootRels ?? [])]),
    "word/document.xml": DOC_XML(),
    "word/styles.xml": `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
    "word/_rels/document.xml.rels": rels([{ id: "rId1", type: T("styles"), target: "styles.xml" }, ...(opts.docRels ?? [])]),
    ...extra,
    "[Content_Types].xml": contentTypes(BASE_DEFAULTS, {
      "word/document.xml": CT_MAIN_DOCX,
      "word/styles.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
      ...(opts.overrides ?? {}),
    }),
  });
}

export interface OpcHata {
  tur: string;
  ayrinti: string;
}

/** Paket tutarlılık hataları (derlem denetimi) + [Content_Types].xml ilk parça mı. */
export function opcHatalari(bytes: Uint8Array, uzanti = ".docx"): OpcHata[] {
  const h = opcDenetle(bytes, { uzanti }) as OpcHata[];
  const ilk = new TextDecoder().decode(bytes.subarray(30, 30 + 19));
  if (ilk !== "[Content_Types].xml") h.push({ tur: "ct-ilk-değil", ayrinti: ilk });
  return h;
}

export const POI = resolve(__dirname, "../../../../test-derlem/poi");
export const poiDosyalari = (re: RegExp) =>
  readdirSync(POI)
    .filter((f) => re.test(f))
    .sort();
export const poi = (f: string) => new Uint8Array(readFileSync(resolve(POI, f)));
