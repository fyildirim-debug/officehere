// DOCX çizimleri: resimler (satır içi ve wp:anchor yerleşimi: metni kaydır, konum, yığın sırası),
// VML resim/filigran ve metin kutuları.
import type { JSONContent } from "@tiptap/core";
import { attr, attrNum, descendants, kid, NS } from "./xml";
import type { PartReader } from "./partReader";
import { CHART_NS, readChartXml } from "./chartXml";

export function readDrawing(pr: PartReader, d: Element, push: (n: JSONContent) => void, extra: JSONContent[]) {
  // Grafik: gömülü grafik parçası (değer önbelleğinden).
  const gd = descendants(d, "graphicData")[0];
  if (gd?.getAttribute("uri") === CHART_NS) {
    const ref = descendants(gd, "chart")[0];
    const rid = ref ? (ref.getAttributeNS(NS.r, "id") ?? attr(ref, "id")) : null;
    const target = rid ? pr.rels.get(rid)?.target : undefined;
    const doc = target ? pr.r.charts.get(target) : undefined;
    const ext = descendants(d, "extent")[0];
    const data = doc ? readChartXml(doc, Math.round((attrNum(ext, "cx") ?? 5486400) / 9525), Math.round((attrNum(ext, "cy") ?? 3200400) / 9525)) : null;
    if (data) push({ type: "chart", attrs: { data } });
    return;
  }
  const blip = descendants(d, "blip")[0];
  const txbx = descendants(d, "txbxContent")[0];
  if (txbx && !blip) {
    pr.boxes.push(textBoxFromDrawing(pr, d, txbx));
    void extra;
    return;
  }
  if (!blip) return;
  const rid = blip.getAttributeNS(NS.r, "embed") ?? attr(blip, "embed");
  const rel = rid ? pr.rels.get(rid) : undefined;
  const src = rel ? pr.r.media.get(rel.target) : undefined;
  if (!src) return;
  const ext = descendants(d, "extent")[0];
  const cx = attrNum(ext, "cx") ?? 1905000;
  const cy = attrNum(ext, "cy") ?? 1428750;
  const docPr = descendants(d, "docPr")[0];
  const xfrm = descendants(d, "xfrm")[0];
  const rot = Math.round((attrNum(xfrm, "rot") ?? 0) / 60000);
  push({
    type: "image",
    attrs: {
      src,
      alt: docPr?.getAttribute("descr") || null,
      width: Math.round(cx / 9525),
      height: Math.round(cy / 9525),
      rotate: rot || 0,
      flipH: xfrm?.getAttribute("flipH") === "1",
      flipV: xfrm?.getAttribute("flipV") === "1",
      crop: cropOf(descendants(d, "srcRect")[0]),
      ...anchorAttrs(d, pr.r.contentWPx, pr.r.marginLPx),
    },
  });
}

/**
 * wp:anchor yerleşimi → metni kaydır, yatay hizalama/konum, yığın sırası. Kaydırmalı resimde yatay uzaklık
 * yaslandığı kenara (sol/sağ), dikey uzaklık paragrafa göre okunur.
 */
export function anchorAttrs(d: Element, contentW = 605, marginL = 94): Record<string, unknown> {
  const anchor = d.localName === "anchor" ? d : descendants(d, "anchor")[0];
  if (!anchor) return {};
  const EMU = 9525;
  const has = (n: string) => !!kid(anchor, n);
  let wrap = "square";
  if (has("wrapNone")) wrap = attr(anchor, "behindDoc") === "1" ? "behind" : "front";
  else if (has("wrapTopAndBottom")) wrap = "topBottom";
  else if (has("wrapTight") || has("wrapThrough")) wrap = "tight";
  const posH = kid(anchor, "positionH");
  const posV = kid(anchor, "positionV");
  const alignH = posH ? kid(posH, "align")?.textContent?.trim() : undefined;
  const offH = posH ? kid(posH, "posOffset")?.textContent : undefined;
  const offV = posV ? kid(posV, "posOffset")?.textContent : undefined;
  const posX = offH ? Math.round(parseInt(offH, 10) / EMU) : 0;
  const posY = offV ? Math.round(parseInt(offV, 10) / EMU) : 0;
  const rel = attrNum(anchor, "relativeHeight");
  const z = rel ? Math.max(-50, Math.min(50, Math.round((rel - 251658240) / 1024))) : 0;
  if (wrap === "behind" || wrap === "front") {
    const floatSide = alignH === "right" || alignH === "outside" ? "right" : alignH === "center" ? "center" : posX > 300 ? "right" : "left";
    return { wrap, floatSide, posX, posY, z };
  }
  const width = (attrNum(kid(anchor, "extent"), "cx") ?? 0) / EMU;
  const fromH = attr(posH, "relativeFrom") ?? "column";
  const vRel = (attr(posV, "relativeFrom") ?? "paragraph") === "paragraph" || attr(posV, "relativeFrom") === "line";
  const y = vRel && !kid(posV, "align") ? Math.max(0, posY) : 0;
  if (alignH || !offH) {
    const floatSide = alignH === "right" || alignH === "outside" ? "right" : alignH === "center" ? "center" : "left";
    return { wrap, floatSide, posX: 0, posY: y, z };
  }
  // Uzaklık sütun/kenar boşluğuna göre (sayfaya göreyse sol kenar boşluğu düşülür).
  const off = fromH === "page" ? posX - marginL : posX;
  if (off + width / 2 > contentW / 2) return { wrap, floatSide: "right", posX: Math.max(0, Math.round(contentW - width - off)), posY: y, z };
  return { wrap, floatSide: "left", posX: Math.max(0, Math.round(off)), posY: y, z };
}

export function readPict(pr: PartReader, pc: Element, push: (n: JSONContent) => void, _extra: JSONContent[]) {
  const tp = descendants(pc, "textpath").find((x) => x.hasAttribute("string"));
  if (tp) {
    const shape = tp.parentElement;
    const style = shape?.getAttribute("style") ?? "";
    const fill = shape?.getAttribute("fillcolor") ?? "#C0C0C0";
    const fontM = /font-family:\s*&?quot;?"?([^";&]+)/.exec(tp.getAttribute("style") ?? "");
    const opacity = descendants(shape ?? pc, "fill")[0]?.getAttribute("opacity");
    pr.watermark = {
      text: tp.getAttribute("string") ?? "",
      color: fill.startsWith("#") ? fill.toUpperCase() : fill === "silver" ? "#C0C0C0" : fill === "gray" ? "#808080" : "#C0C0C0",
      layout: /rotation:\s*(315|-45)/.test(style) ? "diagonal" : "horizontal",
      font: fontM?.[1]?.trim() || "Aptos",
      semitransparent: !!opacity && parseFloat(opacity) < 1,
    };
    return;
  }
  const img = descendants(pc, "imagedata")[0];
  if (img) {
    const rid = img.getAttributeNS(NS.r, "id") ?? attr(img, "id");
    const rel = rid ? pr.rels.get(rid) : undefined;
    const src = rel ? pr.r.media.get(rel.target) : undefined;
    if (!src) return;
    const style = img.parentElement?.getAttribute("style") ?? "";
    const w = /width:\s*([\d.]+)pt/.exec(style);
    const h = /height:\s*([\d.]+)pt/.exec(style);
    push({ type: "image", attrs: { src, width: w ? Math.round(parseFloat(w[1]) * 1.333) : 200, height: h ? Math.round(parseFloat(h[1]) * 1.333) : 150 } });
    return;
  }
  const txbx = descendants(pc, "txbxContent")[0];
  if (txbx) {
    const shape = descendants(pc, "shape")[0] ?? descendants(pc, "rect")[0];
    const st = shape?.getAttribute("style") ?? "";
    const ptv = (k: string, d: number) => {
      const m = new RegExp(`(?:^|;)\\s*${k}:\\s*(-?[\\d.]+)(pt|px|in)?`).exec(st);
      if (!m) return d;
      const v = parseFloat(m[1]);
      return m[2] === "px" ? v : m[2] === "in" ? v * 96 : v / 0.75;
    };
    const color = (v: string | null | undefined) => (v && /^#?[0-9a-f]{6}$/i.test(v) ? `#${v.replace("#", "").toUpperCase()}` : null);
    const filled = shape?.getAttribute("filled");
    const stroked = shape?.getAttribute("stroked");
    pr.boxes.push({
      type: "textBox",
      attrs: {
        x: Math.round(ptv("margin-left", 0)),
        y: Math.round(ptv("margin-top", 0)),
        w: Math.round(ptv("width", 240)),
        h: Math.round(ptv("height", 120)),
        fill: filled === "f" || filled === "false" ? null : (color(shape?.getAttribute("fillcolor")) ?? "#FFFFFF"),
        border: stroked === "f" || stroked === "false" ? null : (color(shape?.getAttribute("strokecolor")) ?? "#000000"),
        borderWidth: parseFloat(shape?.getAttribute("strokeweight") ?? "0.75") || 0.75,
      },
      content: textBoxBlocks(pr, txbx),
    });
  }
}

/** a:srcRect → kırpma oranları. */
function cropOf(el: Element | undefined): { l: number; t: number; r: number; b: number } | null {
  if (!el) return null;
  const v = (n: string) => Math.max(0, Math.min(0.99, (Number(el.getAttribute(n)) || 0) / 100000));
  const c = { l: v("l"), t: v("t"), r: v("r"), b: v("b") };
  return c.l || c.t || c.r || c.b ? c : null;
}

/** Metin kutusu içeriği (yalnız paragraf ve listeler; tablolar paragraflara dönüşür). */
function textBoxBlocks(pr: PartReader, txbx: Element): JSONContent[] {
  const out: JSONContent[] = [];
  const flat = (b: JSONContent) => {
    if (b.type === "paragraph" || b.type === "heading" || b.type === "bulletList" || b.type === "orderedList") out.push(b);
    else if (b.type === "textBox") (b.content ?? []).forEach(flat);
    else if (b.type === "table") for (const row of b.content ?? []) for (const cell of row.content ?? []) (cell.content ?? []).forEach(flat);
  };
  const saved = pr.boxes;
  pr.boxes = [];
  pr.blocks(txbx).forEach(flat);
  pr.boxes = saved;
  return out.length ? out : [{ type: "paragraph" }];
}

/** DrawingML (wp:anchor / wp:inline + wps:wsp) metin kutusu. */
function textBoxFromDrawing(pr: PartReader, d: Element, txbx: Element): JSONContent {
  const EMU = 9525;
  const ext = descendants(d, "extent")[0];
  const posH = descendants(d, "positionH")[0];
  const posV = descendants(d, "positionV")[0];
  const off = (el: Element | undefined) => {
    const o = el ? descendants(el, "posOffset")[0] : undefined;
    return o ? Math.round(parseInt(o.textContent ?? "0", 10) / EMU) : 0;
  };
  let x = off(posH);
  let y = off(posV);
  // Sayfaya göre konumlar kenar boşluğuna göre yaklaşıklanır.
  if (posH && attr(posH, "relativeFrom") === "page") x = Math.max(0, x - 94);
  if (posV && attr(posV, "relativeFrom") === "page") y = Math.max(0, y - 94);
  const spPr = descendants(d, "spPr")[0];
  const ln = spPr ? kid(spPr, "ln") : null;
  const lnColor = ln ? attr(descendants(ln, "srgbClr")[0], "val") : null;
  const noLine = !!ln && !!kid(ln, "noFill");
  const fillEl = spPr ? kid(spPr, "solidFill") : null;
  const fillColor = fillEl ? attr(descendants(fillEl, "srgbClr")[0], "val") : null;
  const noFill = !!spPr && !!kid(spPr, "noFill");
  return {
    type: "textBox",
    attrs: {
      x,
      y,
      w: Math.round((attrNum(ext, "cx") ?? 2286000) / EMU),
      h: Math.round((attrNum(ext, "cy") ?? 1143000) / EMU),
      border: noLine ? null : lnColor ? `#${lnColor.toUpperCase()}` : "#000000",
      borderWidth: ln && attrNum(ln, "w") ? +((attrNum(ln, "w") ?? 9525) / 12700).toFixed(2) : 0.75,
      fill: noFill ? null : fillColor ? `#${fillColor.toUpperCase()}` : "#FFFFFF",
    },
    content: textBoxBlocks(pr, txbx),
  };
}
