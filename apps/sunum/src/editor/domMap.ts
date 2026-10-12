/** Düzenleyici DOM'u ile metin modeli konumları arasında eşleme. */
import type { Pos, Sel } from "../model/text";

function modelLengthOfFragment(frag: DocumentFragment | Element): number {
  let n = 0;
  const walk = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      n += (node as Text).data.length;
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      node.childNodes.forEach(walk);
      return;
    }
    const el = node as Element;
    if (el.tagName === "BR") {
      if (el.getAttribute("data-lb")) n += 1;
      return;
    }
    if (el.getAttribute("data-field")) {
      n += Number(el.getAttribute("data-len") ?? 0);
      return;
    }
    el.childNodes.forEach(walk);
  };
  frag.childNodes.forEach(walk);
  return n;
}

export function paraElements(root: HTMLElement): HTMLElement[] {
  return Array.from(root.querySelectorAll<HTMLElement>(":scope > .tb-p"));
}

/** DOM noktasını model konumuna çevirir. */
export function domToPos(root: HTMLElement, node: Node, offset: number): Pos | null {
  if (!root.contains(node)) return null;
  const paras = paraElements(root);
  if (!paras.length) return { p: 0, o: 0 };
  if (node === root) {
    if (offset >= paras.length) {
      const last = paras.length - 1;
      return { p: last, o: modelLengthOfFragment(paras[last]) };
    }
    return { p: offset, o: 0 };
  }
  const el = node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement;
  const pEl = el?.closest(".tb-p") as HTMLElement | null;
  if (!pEl || !root.contains(pEl)) return null;
  const p = paras.indexOf(pEl);
  const field = el?.closest("[data-field]");
  const range = document.createRange();
  range.setStart(pEl, 0);
  try {
    if (field && pEl.contains(field)) {
      // Alanın içindeyse alanın sonuna yerleştir.
      range.setEndAfter(field);
    } else range.setEnd(node, offset);
  } catch {
    return { p, o: 0 };
  }
  return { p: Math.max(0, p), o: modelLengthOfFragment(range.cloneContents()) };
}

export function domSelToSel(root: HTMLElement): Sel | null {
  const s = window.getSelection();
  if (!s || s.rangeCount === 0 || !s.anchorNode || !s.focusNode) return null;
  const a = domToPos(root, s.anchorNode, s.anchorOffset);
  const f = domToPos(root, s.focusNode, s.focusOffset);
  if (!a || !f) return null;
  return { anchor: a, focus: f };
}

/** Model konumunu DOM noktasına çevirir. */
export function posToDom(root: HTMLElement, pos: Pos): [Node, number] {
  const paras = paraElements(root);
  const pEl = paras[Math.max(0, Math.min(paras.length - 1, pos.p))];
  if (!pEl) return [root, 0];
  let remaining = pos.o;
  let result: [Node, number] | null = null;
  let lastText: Text | null = null;
  const walk = (node: Node): boolean => {
    if (node.nodeType === Node.TEXT_NODE) {
      const t = node as Text;
      lastText = t;
      if (remaining <= t.data.length) {
        result = [t, remaining];
        return true;
      }
      remaining -= t.data.length;
      return false;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return false;
    const el = node as Element;
    if (el.tagName === "BR") {
      if (el.getAttribute("data-lb")) {
        if (remaining === 0) {
          const parent = el.parentNode!;
          result = [parent, Array.prototype.indexOf.call(parent.childNodes, el)];
          return true;
        }
        remaining -= 1;
        if (remaining === 0) {
          const parent = el.parentNode!;
          result = [parent, Array.prototype.indexOf.call(parent.childNodes, el) + 1];
          return true;
        }
      }
      return false;
    }
    if (el.getAttribute("data-field")) {
      const len = Number(el.getAttribute("data-len") ?? 0);
      const parent = el.parentNode!;
      const idx = Array.prototype.indexOf.call(parent.childNodes, el);
      if (remaining === 0) {
        result = [parent, idx];
        return true;
      }
      if (remaining <= len) {
        result = [parent, idx + 1];
        return true;
      }
      remaining -= len;
      return false;
    }
    for (const c of Array.from(el.childNodes)) if (walk(c)) return true;
    return false;
  };
  for (const c of Array.from(pEl.childNodes)) if (walk(c)) break;
  if (result) return result;
  if (lastText) return [lastText, (lastText as Text).data.length];
  const span = pEl.querySelector("span[data-r]");
  if (span) return [span, 0];
  return [pEl, 0];
}

export function applySel(root: HTMLElement, sel: Sel) {
  const s = window.getSelection();
  if (!s) return;
  const [an, ao] = posToDom(root, sel.anchor);
  const [fn, fo] = posToDom(root, sel.focus);
  try {
    s.setBaseAndExtent(an, ao, fn, fo);
  } catch {
    // yok say
  }
}

/** Ekran noktasındaki model konumu. */
export function posFromPoint(root: HTMLElement, x: number, y: number): Pos | null {
  const doc = document as Document & {
    caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
    caretRangeFromPoint?: (x: number, y: number) => Range | null;
  };
  if (doc.caretPositionFromPoint) {
    const cp = doc.caretPositionFromPoint(x, y);
    if (cp && root.contains(cp.offsetNode)) return domToPos(root, cp.offsetNode, cp.offset);
  }
  if (doc.caretRangeFromPoint) {
    const r = doc.caretRangeFromPoint(x, y);
    if (r && root.contains(r.startContainer)) return domToPos(root, r.startContainer, r.startOffset);
  }
  return null;
}
