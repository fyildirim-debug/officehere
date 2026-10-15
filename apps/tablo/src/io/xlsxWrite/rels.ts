/** İlişki (.rels) parçaları: kimlikler eklenme sırasıyla rId1, rId2… */
import { esc, XML_DECL } from "./xml";

const OD = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

export const REL = {
  officeDocument: OD + "officeDocument",
  coreProperties: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  extendedProperties: OD + "extended-properties",
  styles: OD + "styles",
  theme: OD + "theme",
  sharedStrings: OD + "sharedStrings",
  worksheet: OD + "worksheet",
  hyperlink: OD + "hyperlink",
  comments: OD + "comments",
  vmlDrawing: OD + "vmlDrawing",
  drawing: OD + "drawing",
  image: OD + "image",
  table: OD + "table",
  sheetMetadata: OD + "sheetMetadata",
  featurePropertyBag: "http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag",
};

export interface Rel {
  id: string;
  type: string;
  target: string;
  external?: boolean;
}

export class Rels {
  readonly list: Rel[] = [];

  add(type: string, target: string, external = false): string {
    const id = `rId${this.list.length + 1}`;
    this.list.push({ id, type, target, external });
    return id;
  }

  get size(): number {
    return this.list.length;
  }

  xml(): string {
    return relsXml(this.list);
  }
}

export function relsXml(list: readonly Rel[]): string {
  const items = list.map((r) => `<Relationship Id="${esc(r.id)}" Type="${esc(r.type)}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ""}/>`);
  return `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.join("")}</Relationships>`;
}

/** Var olan .rels metnine ilişki ekler (kimlik mevcutların en büyüğünden sonra). */
export function appendRel(xml: string, type: string, target: string): string {
  let n = 1;
  for (const m of xml.matchAll(/Id="rId(\d+)"/g)) n = Math.max(n, Number(m[1]) + 1);
  return xml.replace("</Relationships>", `<Relationship Id="rId${n}" Type="${esc(type)}" Target="${esc(target)}"/></Relationships>`);
}
