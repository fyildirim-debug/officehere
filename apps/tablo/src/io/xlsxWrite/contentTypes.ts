/** [Content_Types].xml: varsayılan uzantılar ve parça başına geçersiz kılmalar. */
import { esc, XML_DECL } from "./xml";

const OFD = "application/vnd.openxmlformats-officedocument.";

export const CT = {
  rels: "application/vnd.openxmlformats-package.relationships+xml",
  xml: "application/xml",
  workbook: OFD + "spreadsheetml.sheet.main+xml",
  worksheet: OFD + "spreadsheetml.worksheet+xml",
  theme: OFD + "theme+xml",
  styles: OFD + "spreadsheetml.styles+xml",
  sharedStrings: OFD + "spreadsheetml.sharedStrings+xml",
  table: OFD + "spreadsheetml.table+xml",
  drawing: OFD + "drawing+xml",
  vml: OFD + "vmlDrawing",
  comments: OFD + "spreadsheetml.comments+xml",
  core: "application/vnd.openxmlformats-package.core-properties+xml",
  app: OFD + "extended-properties+xml",
  sheetMetadata: OFD + "spreadsheetml.sheetMetadata+xml",
  featurePropertyBag: "application/vnd.ms-excel.featurepropertybag+xml",
};

export interface ContentTypesInput {
  /** Resim uzantıları (ilk görülme sırasıyla). */
  media: string[];
  sheets: string[];
  sharedStrings: boolean;
  tables: string[];
  drawings: string[];
  comments: string[];
}

/** Bilinen parçaların içerik türleri (sıra: varsayılanlar, kitap, sayfalar, tema, stiller…). */
export function contentTypesXml(o: ContentTypesInput): string {
  const out: string[] = [];
  const def = (ext: string, type: string) => out.push(`<Default Extension="${esc(ext)}" ContentType="${esc(type)}"/>`);
  const over = (part: string, type: string) => out.push(`<Override PartName="${esc(part)}" ContentType="${esc(type)}"/>`);
  for (const ext of o.media) def(ext, `image/${ext}`);
  def("rels", CT.rels);
  def("xml", CT.xml);
  over("/xl/workbook.xml", CT.workbook);
  for (const s of o.sheets) over(`/${s}`, CT.worksheet);
  over("/xl/theme/theme1.xml", CT.theme);
  over("/xl/styles.xml", CT.styles);
  if (o.sharedStrings) over("/xl/sharedStrings.xml", CT.sharedStrings);
  for (const t of o.tables) over(`/${t}`, CT.table);
  for (const d of o.drawings) over(`/${d}`, CT.drawing);
  def("vml", CT.vml);
  for (const c of o.comments) over(`/${c}`, CT.comments);
  over("/docProps/core.xml", CT.core);
  over("/docProps/app.xml", CT.app);
  return `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${out.join("")}</Types>`;
}

/** Var olan [Content_Types].xml metnine parça türü ekler (zaten varsa dokunmaz). */
export function addOverride(xml: string, part: string, type: string): string {
  if (xml.includes(`PartName="${part}"`)) return xml;
  return xml.replace("</Types>", `<Override PartName="${esc(part)}" ContentType="${esc(type)}"/></Types>`);
}
