/** Köprü komutları: Ekle > Bağlantı (Ctrl+K), Bağlantıyı Düzenle/Aç/Kaldır. */
import { makeTextBox } from "../model/factory";
import { linkSlideIndex, parseLink } from "../model/links";
import { findShape, updateSlideShapes } from "../model/ops";
import { isCollapsed, orderSel, paraLen, runPropsInRange, type Pos } from "../model/text";
import type { RunProps, TextBody } from "../model/types";
import type { Actions } from "./actions";

export interface LinkContext {
  /** text: düzenlenen metnin seçimi; shape: seçili şeklin kendisi; new: yeni metin kutusu. */
  mode: "text" | "shape" | "new";
  link?: string;
  /** Görüntülenecek metin. */
  text: string;
  textEditable: boolean;
}

function editingBody(a: Actions): TextBody | undefined {
  const ed = a.s.editing;
  if (!ed) return undefined;
  const sh = findShape(a.slide()?.shapes ?? [], ed.shapeId);
  if (!sh) return undefined;
  if (sh.kind === "sp") return sh.text;
  if (sh.kind === "table" && ed.cell) return sh.rows[ed.cell[0]]?.cells[ed.cell[1]]?.text;
  return undefined;
}

/** İmlecin bulunduğu köprü çalıştırmasının aralığı. */
function linkRunAt(body: TextBody, pos: Pos): { from: Pos; to: Pos; link: string } | null {
  const p = body.paragraphs[pos.p];
  if (!p) return null;
  let acc = 0;
  for (const r of p.runs) {
    const s = acc;
    acc += r.text.length;
    if (r.props.link && pos.o >= s && pos.o <= acc) return { from: { p: pos.p, o: s }, to: { p: pos.p, o: acc }, link: r.props.link };
  }
  return null;
}

/** Bağlantı iletişim kutusu için bağlam (gerekirse seçimi köprünün tamamına genişletir). */
export function linkContext(a: Actions): LinkContext {
  const h = a.textHandle;
  const body = editingBody(a);
  if (a.s.editing && h && body) {
    let sel = h.selection();
    if (isCollapsed(sel)) {
      const hit = linkRunAt(body, sel.anchor);
      if (hit) {
        sel = { anchor: hit.from, focus: hit.to };
        h.setSelection(sel);
      }
    }
    const [from, to] = orderSel(sel);
    const text = h.selectedText();
    const props = runPropsInRange(body, from, to);
    const link = props.find((p) => p.link)?.link ?? (isCollapsed(sel) ? undefined : linkRunAt(body, from)?.link);
    return { mode: "text", link, text, textEditable: !text.includes("\n") && from.p === to.p && to.o <= paraLen(body.paragraphs[to.p]) };
  }
  const shapes = a.selectedShapes();
  if (shapes.length === 1) return { mode: "shape", link: shapes[0].link, text: "<<Belgedeki seçim>>", textEditable: false };
  return { mode: "new", text: "", textEditable: true };
}

function stripLink(p: RunProps): RunProps {
  if (!p.link) return p;
  const { link: _l, ...rest } = p;
  void _l;
  return rest;
}

/** Köprüyü uygular (link null ise kaldırır). */
export function applyLink(a: Actions, c: LinkContext, link: string | null, text: string) {
  const h = a.textHandle;
  if (c.mode === "text" && h && a.s.editing) {
    if (!link) h.applyRun((p) => stripLink(p));
    else if (!c.text || (c.textEditable && text && text !== c.text)) h.insertRun(text || link.replace(/^mailto:/, ""), { link });
    else h.applyRun((p) => ({ ...p, link }));
    h.focus();
    a.store.set((s) => ({ textTick: s.textTick + 1 }));
    return;
  }
  if (c.mode === "shape") {
    a.updateShapes((s) => {
      const n = { ...s };
      if (link) n.link = link;
      else delete n.link;
      return n;
    });
    return;
  }
  if (!link) return;
  // Hiçbir şey seçili değilse köprü metni yeni bir metin kutusuna eklenir.
  const p = a.pres;
  const label = text || link.replace(/^mailto:/, "");
  const w = Math.min(p.width * 0.5, Math.max(160, label.length * 14));
  const tb = makeTextBox((p.width - w) / 2, p.height / 2 - 24, w, 48, "");
  tb.name = `Metin Kutusu ${a.nextShapeNumber()}`;
  tb.text = { ...tb.text!, paragraphs: [{ runs: [{ text: label, props: { link } }], level: 0, props: {} }] };
  const i = a.s.slideIdx;
  if (!a.slide()) return;
  a.store.commit((pp) => updateSlideShapes(pp, i, (list) => [...list, tb]), { sel: [tb.id], editing: null });
}

/** Seçili şeklin (yoksa düzenlenen metnin) köprüsü. */
export function currentLink(a: Actions): string | undefined {
  const body = editingBody(a);
  const h = a.textHandle;
  if (a.s.editing && h && body) {
    const sel = h.selection();
    const [from, to] = orderSel(sel);
    return runPropsInRange(body, from, to).find((p) => p.link)?.link ?? linkRunAt(body, from)?.link;
  }
  const shapes = a.selectedShapes();
  if (shapes.length !== 1) return undefined;
  const s = shapes[0];
  if (s.link) return s.link;
  if (s.kind === "sp" && s.text) for (const p of s.text.paragraphs) for (const r of p.runs) if (r.props.link) return r.props.link;
  return undefined;
}

/** Düzenleme görünümünde köprüyü izler (Bağlantıyı Aç / Ctrl+tıklama). */
export function followLink(a: Actions, link: string) {
  const t = parseLink(link);
  if (t.kind === "url") {
    window.open(t.url, "_blank", "noopener,noreferrer");
    return;
  }
  const i = linkSlideIndex(link, a.pres, a.s.slideIdx);
  if (i !== null && i >= 0) a.goToSlide(i);
}

/** Seçimdeki köprüleri kaldırır. */
export function removeLink(a: Actions) {
  const h = a.textHandle;
  const body = editingBody(a);
  if (a.s.editing && h && body) {
    const sel = h.selection();
    if (isCollapsed(sel)) {
      const hit = linkRunAt(body, sel.anchor);
      if (hit) h.setSelection({ anchor: hit.from, focus: hit.to });
    }
    h.applyRun((p) => stripLink(p));
    return;
  }
  a.updateShapes((s) => {
    if (s.link) {
      const n = { ...s };
      delete n.link;
      return n;
    }
    if (s.kind === "sp" && s.text) return { ...s, text: { ...s.text, paragraphs: s.text.paragraphs.map((p) => ({ ...p, runs: p.runs.map((r) => (r.props.link ? { ...r, props: stripLink(r.props) } : r)) })) } };
    return s;
  });
}
