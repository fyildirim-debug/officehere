// Resim/şekil yerleşimi (Düzen › Yerleştir, Resim Biçimi): Metni Kaydır (satır içi, kare, sıkı, üst ve alt,
// metnin arkasında, metnin önünde), Konum önayarları, Öne Getir / Arkaya Gönder. Satır içi olmayan resim
// DOCX'te wp:anchor olarak yazılır. Kaydırmalı tüm resimler sürüklenerek taşınır: metnin arkasında/önündeki
// resim serbestçe konumlanır; kare/sıkı/üst-alt resmin çapası bırakıldığı paragrafa taşınır ve yatay/dikey
// uzaklığı güncellenir (Word gibi).
import type { Editor } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { NodeSelection, type Transaction } from "@tiptap/pm/state";
import { closeHistory } from "@tiptap/pm/history";
import type { EditorView } from "@tiptap/pm/view";

export type WrapMode = "inline" | "square" | "tight" | "topBottom" | "behind" | "front";

export const WRAP_LABEL: Record<WrapMode, string> = {
  inline: "Metinle Aynı Hizada",
  square: "Kare",
  tight: "Sıkı",
  topBottom: "Üst ve Alt",
  behind: "Metnin Arkasında",
  front: "Metnin Önünde",
};

export interface WrapAttrs {
  wrap: WrapMode;
  /** Kare/sıkı/üst-alt: yatay yerleşim. */
  floatSide: "left" | "right" | "center";
  /**
   * Bağlı paragrafa göre konum (px). Metnin arkasında/önünde: sol üst köşe. Kare/sıkı/üst-alt: posX
   * yaslandığı kenardan (sol ya da sağ) uzaklık, posY paragrafın üstünden uzaklık.
   */
  posX: number;
  posY: number;
  /** Yığın sırası (Öne Getir / Arkaya Gönder). */
  z: number;
}

/** Resim düğümüne eklenen yerleşim öznitelikleri. */
export const wrapAttributes = {
  wrap: { default: "inline", parseHTML: (el: HTMLElement) => el.getAttribute("data-wrap") ?? "inline", renderHTML: (a: Record<string, unknown>) => (a.wrap && a.wrap !== "inline" ? { "data-wrap": a.wrap } : {}) },
  floatSide: { default: "left", parseHTML: (el: HTMLElement) => el.getAttribute("data-float") ?? "left", renderHTML: (a: Record<string, unknown>) => (a.wrap && a.wrap !== "inline" ? { "data-float": a.floatSide } : {}) },
  posX: { default: 0, parseHTML: (el: HTMLElement) => Number(el.getAttribute("data-x") ?? 0), renderHTML: (a: Record<string, unknown>) => (a.posX ? { "data-x": String(a.posX) } : {}) },
  posY: { default: 0, parseHTML: (el: HTMLElement) => Number(el.getAttribute("data-y") ?? 0), renderHTML: (a: Record<string, unknown>) => (a.posY ? { "data-y": String(a.posY) } : {}) },
  z: { default: 0, parseHTML: (el: HTMLElement) => Number(el.getAttribute("data-z") ?? 0), renderHTML: (a: Record<string, unknown>) => (a.z ? { "data-z": String(a.z) } : {}) },
};

/** Yerleşimin CSS'i (düzenleyicide resim kapsayıcısına, dışa aktarmada resme uygulanır). */
export function wrapStyle(a: Partial<WrapAttrs>): Partial<CSSStyleDeclaration> {
  const s: Partial<CSSStyleDeclaration> = { float: "", position: "", left: "", top: "", zIndex: "", margin: "", display: "", justifyContent: "" };
  switch (a.wrap) {
    case "square":
    case "tight": {
      const gap = a.wrap === "tight" ? "6px" : "12px";
      const x = Math.max(0, a.posX ?? 0);
      const y = Math.max(0, a.posY ?? 0);
      if (a.floatSide === "center") Object.assign(s, { display: "flex", justifyContent: "center", margin: `${6 + y}px auto 6px` });
      else Object.assign(s, { float: a.floatSide === "right" ? "right" : "left", margin: a.floatSide === "right" ? `${y}px ${x}px 6px ${gap}` : `${y}px ${gap} 6px ${x}px` });
      break;
    }
    case "topBottom": {
      const x = Math.max(0, a.posX ?? 0);
      const y = 6 + Math.max(0, a.posY ?? 0);
      // Düzenleyicide kapsayıcı tam genişliktedir (yaslama justify-content ile); dışa aktarmada resmin kendisi (otomatik kenar boşluğu).
      const justifyContent = a.floatSide === "right" ? "flex-end" : a.floatSide === "center" ? "center" : "flex-start";
      Object.assign(s, { display: "flex", justifyContent, margin: a.floatSide === "right" ? `${y}px ${x}px 6px auto` : a.floatSide === "center" ? `${y}px auto 6px` : `${y}px auto 6px ${x}px` });
      break;
    }
    case "behind":
    case "front":
      Object.assign(s, { position: "absolute", left: `${a.posX ?? 0}px`, top: `${a.posY ?? 0}px`, zIndex: String(a.wrap === "behind" ? -1 : 3 + (a.z ?? 0)) });
      break;
  }
  return s;
}

/** Seçili resim (NodeSelection) ve konumu. */
export function selectedImage(editor: Editor): { node: PMNode; pos: number } | null {
  const sel = editor.state.selection;
  return sel instanceof NodeSelection && sel.node.type.name === "image" ? { node: sel.node, pos: sel.from } : null;
}

function update(editor: Editor, pos: number, attrs: Record<string, unknown>) {
  const node = editor.state.doc.nodeAt(pos);
  if (!node) return;
  const tr = editor.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...attrs });
  tr.setSelection(NodeSelection.create(tr.doc, pos));
  editor.view.dispatch(tr);
}

/** Metni Kaydır. */
export function setWrap(editor: Editor, wrap: WrapMode) {
  const img = selectedImage(editor);
  if (!img) return;
  const patch: Record<string, unknown> = { wrap };
  // Arkaya/öne alınan resim bulunduğu yerde kalır: konum, paragrafa göre şimdiki yeri olur.
  if ((wrap === "behind" || wrap === "front") && img.node.attrs.wrap !== "behind" && img.node.attrs.wrap !== "front") {
    const el = editor.view.nodeDOM(img.pos) as HTMLElement | null;
    const para = el?.closest("p, h1, h2, h3, h4, h5, h6") as HTMLElement | null;
    if (el && para) {
      const scale = zoomOf(editor.view);
      const a = el.getBoundingClientRect();
      const b = para.getBoundingClientRect();
      patch.posX = Math.round((a.left - b.left) / scale);
      patch.posY = Math.round((a.top - b.top) / scale);
    }
  }
  // Serbest konumdan kaydırmalıya: sol üst köşe korunur (soldan uzaklık ve paragraf üstünden uzaklık).
  const wasFree = img.node.attrs.wrap === "behind" || img.node.attrs.wrap === "front";
  if (wasFree && (wrap === "square" || wrap === "tight" || wrap === "topBottom")) Object.assign(patch, { floatSide: "left", posX: Math.max(0, img.node.attrs.posX as number), posY: Math.max(0, img.node.attrs.posY as number) });
  if (wrap === "inline") Object.assign(patch, { posX: 0, posY: 0 });
  update(editor, img.pos, patch);
}

/** Öne Getir / Arkaya Gönder (bir düzey ya da en öne/en arkaya). */
export function arrange(editor: Editor, dir: "forward" | "front" | "backward" | "back") {
  const img = selectedImage(editor);
  if (!img) return;
  const zs: number[] = [];
  editor.state.doc.descendants((n) => {
    if (n.type.name === "image" && n.attrs.wrap !== "inline") zs.push(n.attrs.z as number);
    return true;
  });
  const z = img.node.attrs.z as number;
  let next = z;
  if (dir === "front") next = Math.max(z, ...zs) + 1;
  else if (dir === "back") next = Math.min(z, ...zs) - 1;
  else next = z + (dir === "forward" ? 1 : -1);
  const patch: Record<string, unknown> = { z: next };
  // Word'de en arkaya gönderilen resim metnin arkasına, en öne getirilen metnin önüne geçer.
  if (img.node.attrs.wrap === "inline") patch.wrap = dir === "back" || dir === "backward" ? "behind" : "front";
  update(editor, img.pos, patch);
}

export interface PositionPreset {
  id: string;
  label: string;
  v: "top" | "middle" | "bottom";
  h: "left" | "center" | "right";
}

const V_LABEL = { top: "Üst", middle: "Orta", bottom: "Alt" };
const H_LABEL = { left: "Sol", center: "Orta", right: "Sağ" };

export const POSITION_PRESETS: PositionPreset[] = (["top", "middle", "bottom"] as const).flatMap((v) =>
  (["left", "center", "right"] as const).map((h) => ({ id: `${v}-${h}`, label: `${V_LABEL[v]} ${H_LABEL[h]}, Kare Metin Kaydırma ile`, v, h })),
);

/**
 * Konum: resim imlecin bulunduğu sayfanın üst/orta/alt kısmındaki paragrafa taşınır ve kare metin
 * kaydırmayla sola/ortaya/sağa yerleştirilir.
 */
export function setPosition(editor: Editor, preset: PositionPreset, page: { top: number; h: number; mt: number; mb: number }, pagesTop: number, scale: number) {
  const img = selectedImage(editor);
  if (!img) return;
  const view = editor.view;
  const contentTop = pagesTop + (page.top + page.mt) * scale;
  const contentBottom = pagesTop + (page.top + page.h - page.mb) * scale;
  const y = preset.v === "top" ? contentTop + 4 : preset.v === "middle" ? (contentTop + contentBottom) / 2 : contentBottom - ((img.node.attrs.height as number) ?? 100) * scale - 4;
  const dom = view.dom.getBoundingClientRect();
  const hit = view.posAtCoords({ left: dom.left + dom.width / 2, top: y });
  let tr = view.state.tr;
  const node = img.node;
  let at = img.pos;
  if (hit) {
    const $t = view.state.doc.resolve(hit.pos);
    if ($t.parent.isTextblock) {
      const target = $t.start();
      if (target !== img.pos && target !== img.pos + 1) {
        tr = tr.delete(img.pos, img.pos + node.nodeSize);
        at = tr.mapping.map(target);
        tr.insert(at, node);
      }
    }
  }
  tr.setNodeMarkup(at, undefined, { ...node.attrs, wrap: preset.h === "center" ? "topBottom" : "square", floatSide: preset.h, posX: 0, posY: 0 });
  tr.setSelection(NodeSelection.create(tr.doc, at));
  view.dispatch(tr.scrollIntoView());
}

function zoomOf(view: EditorView): number {
  const host = view.dom.closest(".bl-canvas") as HTMLElement | null;
  return host ? parseFloat(getComputedStyle(host).zoom) || 1 : 1;
}

/** Kaydırmalı resmi sürükleyerek taşıma (resim kapsayıcısında). */
export function enableFloatingDrag(container: HTMLElement, view: EditorView, getPos: () => number | undefined) {
  container.addEventListener("mousedown", (e) => {
    const pos = getPos();
    if (pos === undefined || e.button !== 0) return;
    if ((e.target as HTMLElement).closest("[data-resize-handle]")) return;
    const node = view.state.doc.nodeAt(pos);
    const wrap = node?.attrs.wrap as WrapMode | undefined;
    if (!node || !wrap || wrap === "inline") return;
    e.preventDefault();
    view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)));
    if (wrap === "behind" || wrap === "front") dragFree(e, container, view, getPos, node);
    else dragAnchored(e, container, view, getPos);
  });
}

/** Metnin arkasında/önünde: paragrafa göre serbest konum. */
function dragFree(e: MouseEvent, container: HTMLElement, view: EditorView, getPos: () => number | undefined, node: PMNode) {
  const scale = zoomOf(view);
  const sx = e.clientX;
  const sy = e.clientY;
  const x0 = node.attrs.posX as number;
  const y0 = node.attrs.posY as number;
  let moved = false;
  const move = (ev: MouseEvent) => {
    const dx = (ev.clientX - sx) / scale;
    const dy = (ev.clientY - sy) / scale;
    if (!moved && Math.hypot(dx, dy) < 3) return;
    moved = true;
    container.style.left = `${Math.round(x0 + dx)}px`;
    container.style.top = `${Math.round(y0 + dy)}px`;
  };
  const up = (ev: MouseEvent) => {
    window.removeEventListener("mousemove", move);
    window.removeEventListener("mouseup", up);
    if (!moved) return;
    const p = getPos();
    const cur = p !== undefined ? view.state.doc.nodeAt(p) : null;
    if (p === undefined || !cur) return;
    const tr = view.state.tr.setNodeMarkup(p, undefined, { ...cur.attrs, posX: Math.round(x0 + (ev.clientX - sx) / scale), posY: Math.round(y0 + (ev.clientY - sy) / scale) });
    tr.setSelection(NodeSelection.create(tr.doc, p));
    closeHistory(tr);
    view.dispatch(tr);
    view.focus();
  };
  window.addEventListener("mousemove", move);
  window.addEventListener("mouseup", up);
}

/** Kaydırma yerleşimi: resmin yeni sol üst köşesine göre yaslanma kenarı ve uzaklıklar. */
export function anchoredPlacement(wrap: WrapMode, img: { left: number; top: number; width: number }, para: { left: number; top: number; width: number }) {
  const offL = img.left - para.left;
  const offR = para.left + para.width - (img.left + img.width);
  const posY = Math.max(0, Math.round(img.top - para.top));
  // Üst ve Alt: ortaya yakın bırakılan resim ortalanır.
  if (wrap === "topBottom" && Math.abs(offL - offR) < 24) return { floatSide: "center" as const, posX: 0, posY };
  if (offL <= offR) return { floatSide: "left" as const, posX: Math.max(0, Math.round(offL)), posY };
  return { floatSide: "right" as const, posX: Math.max(0, Math.round(offR)), posY };
}

/** Kare/sıkı/üst-alt: resim bırakıldığı paragrafa bağlanır; kenar ve uzaklıklar yeni konumdan hesaplanır. */
function dragAnchored(e: MouseEvent, container: HTMLElement, view: EditorView, getPos: () => number | undefined) {
  const scale = zoomOf(view);
  const sx = e.clientX;
  const sy = e.clientY;
  let moved = false;
  const reset = () => {
    container.style.transform = "";
    container.style.pointerEvents = "";
    container.classList.remove("is-dragging");
  };
  const move = (ev: MouseEvent) => {
    const dx = (ev.clientX - sx) / scale;
    const dy = (ev.clientY - sy) / scale;
    if (!moved && Math.hypot(dx, dy) < 3) return;
    moved = true;
    // Sürükleme sırasında resim gölge gibi izler; bırakma noktası altındaki metinden okunur.
    container.style.pointerEvents = "none";
    container.classList.add("is-dragging");
    container.style.transform = `translate(${dx}px, ${dy}px)`;
  };
  const up = (ev: MouseEvent) => {
    window.removeEventListener("mousemove", move);
    window.removeEventListener("mouseup", up);
    if (!moved) return reset();
    const r = container.getBoundingClientRect();
    const hit = view.posAtCoords({ left: Math.max(r.left, 1), top: r.top + 2 }) ?? view.posAtCoords({ left: ev.clientX, top: ev.clientY });
    reset();
    const p = getPos();
    const cur = p !== undefined ? view.state.doc.nodeAt(p) : null;
    if (p === undefined || !cur || !hit) return;
    // Hedef paragraf: bırakılan noktadaki metin bloğu (yoksa resmin şimdiki paragrafı).
    let $t = view.state.doc.resolve(hit.inside >= 0 ? hit.inside : hit.pos);
    if (!$t.parent.isTextblock) $t = view.state.doc.resolve(hit.pos);
    if (!$t.parent.isTextblock) $t = view.state.doc.resolve(p);
    const paraPos = $t.before();
    const paraEl = view.nodeDOM(paraPos) as HTMLElement | null;
    if (!paraEl) return;
    const pr = paraEl.getBoundingClientRect();
    const place = anchoredPlacement(cur.attrs.wrap as WrapMode, { left: r.left / scale, top: r.top / scale, width: r.width / scale }, { left: pr.left / scale, top: pr.top / scale, width: pr.width / scale });
    let tr = view.state.tr;
    let at = p;
    const start = $t.start();
    const $img = view.state.doc.resolve(p);
    // Çapa başka paragrafa taşınır (Word gibi resim paragrafın başına bağlanır).
    if ($img.parent !== $t.parent || $img.start() !== start) {
      tr = tr.delete(p, p + cur.nodeSize);
      at = tr.mapping.map(start);
      tr.insert(at, cur);
    }
    tr.setNodeMarkup(at, undefined, { ...cur.attrs, ...place });
    tr.setSelection(NodeSelection.create(tr.doc, at));
    // Sürükleme ayrı bir geri alma adımıdır.
    closeHistory(tr);
    view.dispatch(tr);
    // Çapa taşınınca metin yeniden akar: resmin gerçek yeri ölçülüp uzaklıklar düzeltilir (aynı geri al adımı).
    correctPlacement(view, at, r, scale, tr);
    view.focus();
  };
  window.addEventListener("mousemove", move);
  window.addEventListener("mouseup", up);
}

function correctPlacement(view: EditorView, at: number, want: DOMRect, scale: number, prev: Transaction) {
  const node = view.state.doc.nodeAt(at);
  const el = (view.nodeDOM(at) as HTMLElement | null)?.querySelector("img");
  if (!node || !el) return;
  const got = el.getBoundingClientRect();
  const dy = (want.top - got.top) / scale;
  const side = node.attrs.floatSide as string;
  const dx = side === "center" ? 0 : ((side === "right" ? got.left - want.left : want.left - got.left) / scale);
  if (Math.abs(dy) < 1 && Math.abs(dx) < 1) return;
  const attrs = { ...node.attrs, posX: Math.max(0, Math.round((node.attrs.posX as number) + dx)), posY: Math.max(0, Math.round((node.attrs.posY as number) + dy)) };
  const tr = view.state.tr.setNodeMarkup(at, undefined, attrs);
  tr.setSelection(NodeSelection.create(tr.doc, at));
  // Önceki işlemle aynı geri alma adımında birleşir.
  tr.setMeta("appendedTransaction", prev);
  view.dispatch(tr);
}

const IMG_RESET = { float: "", position: "", left: "", top: "", zIndex: "", margin: "", display: "", justifyContent: "" };

/** Yerleşimi resim kapsayıcısına uygular (resmin kendi satır içi yerleşim stilleri temizlenir). */
export function applyWrap(container: HTMLElement, img: HTMLElement, attrs: Partial<WrapAttrs>) {
  Object.assign(img.style, IMG_RESET);
  Object.assign(container.style, wrapStyle(attrs));
  // Satır içi resimde kapsayıcı satır içi esnek kutu kalmalı; blok olursa satır genişliğine yayılır ve tutamaçlar resmin dışına düşer.
  if (!container.style.display) container.style.display = "inline-flex";
  container.dataset.wrap = attrs.wrap ?? "inline";
}
