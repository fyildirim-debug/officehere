/**
 * Resimler: sayfa başına drawingN.xml (tek hücreye bağlı çapalar) ve xl/media/imageN.*. Grafikler
 * ve taşınan şekiller aynı çizime sonradan eklenir (writeChartsInto, restorePassthrough).
 */
import type { SheetModel } from "../../model/workbook";
import { widthPxToChars, pxToPt } from "../../model/axis";
import { Rels, REL } from "./rels";
import { XML_DECL } from "./xml";

export interface MediaItem {
  path: string;
  ext: string;
  data: Uint8Array;
}

interface AnchorOut {
  col: number;
  colOff: number;
  row: number;
  rowOff: number;
  cx: number;
  cy: number;
  rId: string;
}

function base64Bytes(b64: string): Uint8Array | null {
  try {
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  } catch {
    return null;
  }
}

/** Sütunun çapa birimi (genişlik × 10000; özel genişlik yoksa 640000) ve satırınki (yükseklik pt × 10000; yoksa 180000). */
function colUnit(sheet: SheetModel, styled: Set<number>, c: number): number {
  const inSet = sheet.cols.isCustom(c) || sheet.cols.isHidden(c) || styled.has(c) || !!sheet.cols.level(c);
  if (!inSet) return 640000;
  const w = widthPxToChars(sheet.cols.isCustom(c) ? sheet.cols.rawSize(c) : sheet.cols.defaultSize);
  return w !== 9 ? Math.floor(w * 10000) : 640000;
}

function rowUnit(sheet: SheetModel, r: number): number {
  const h = sheet.rows.isCustom(r) ? pxToPt(sheet.rows.rawSize(r)) : 0;
  return h ? Math.floor(h * 10000) : 180000;
}

export interface DrawingOut {
  path: string;
  xml: string;
  rels: Rels;
}

/**
 * Sayfanın resim çizimi. `media` kitap genelindeki resim listesidir (sırayla image1, image2…);
 * çözülemeyen resim atlanır.
 */
export function buildDrawing(sheet: SheetModel, n: number, media: MediaItem[]): DrawingOut | null {
  const rels = new Rels();
  const anchors: AnchorOut[] = [];
  const styled = new Set(sheet.colStyles.keys());
  for (const img of sheet.images) {
    const m = /^data:image\/(png|jpe?g|gif);base64,(.*)$/.exec(img.src);
    if (!m) continue;
    const ext = m[1] === "jpg" ? "jpeg" : m[1];
    const data = base64Bytes(img.src.slice(img.src.indexOf(",") + 1));
    if (!data) continue;
    const path = `xl/media/image${media.length + 1}.${ext}`;
    media.push({ path, ext, data });
    const vc = img.anchor.col + img.anchor.dx / Math.max(1, sheet.cols.size(img.anchor.col));
    const vr = img.anchor.row + img.anchor.dy / Math.max(1, sheet.rows.size(img.anchor.row));
    const col = Math.floor(vc);
    const row = Math.floor(vr);
    anchors.push({
      col,
      colOff: Math.floor((vc - col) * colUnit(sheet, styled, col)),
      row,
      rowOff: Math.floor((vr - row) * rowUnit(sheet, row)),
      cx: Math.floor(img.width * 9525),
      cy: Math.floor(img.height * 9525),
      rId: rels.add(REL.image, `../media/${path.slice("xl/media/".length)}`),
    });
  }
  if (!anchors.length) return null;
  const body = anchors
    .map(
      (a, i) =>
        `<xdr:oneCellAnchor><xdr:from><xdr:col>${a.col}</xdr:col><xdr:colOff>${a.colOff}</xdr:colOff><xdr:row>${a.row}</xdr:row><xdr:rowOff>${a.rowOff}</xdr:rowOff></xdr:from><xdr:ext cx="${a.cx}" cy="${a.cy}"/>` +
        `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 1}" name="Picture ${i + 1}"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{00000000-0008-0000-0000-000002000000}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
        `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${a.rId}" cstate="print"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
        '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>',
    )
    .join("");
  const xml = `${XML_DECL}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${body}</xdr:wsDr>`;
  return { path: `xl/drawings/drawing${n}.xml`, xml, rels };
}
