// Satır numaraları (Düzen › Satır Numaraları): bölüm ayarına göre satırların numaralanması (saf mantık)
// ve düzenleyici DOM'undan satır kutularının ölçülmesi. Tablolar, metin kutuları, üst/alt bilgi ve
// dipnotlar numaralanmaz (Word gibi); "Geçerli Paragrafta Gizle" paragrafı sayılmaz.
import type { EditorView } from "@tiptap/pm/view";
import type { LineNumbering } from "../model/document";

export interface LineBox {
  /** Sayfa kapsayıcısı koordinatı (ölçeksiz). */
  top: number;
  height: number;
  left: number;
  section: number;
  page: number;
  suppress: boolean;
}

export interface NumberedLine {
  line: LineBox;
  n: number;
}

/** Satırları bölüm ayarına göre numaralar; gösterilecek (countBy katı) satırları döner. */
export function numberLines(lines: LineBox[], cfg: (section: number) => LineNumbering | null | undefined): NumberedLine[] {
  const out: NumberedLine[] = [];
  let n = 0;
  let lastSection = -1;
  let lastPage = -1;
  let started = false;
  for (const l of lines) {
    const c = cfg(l.section);
    if (!c) continue;
    const start = Math.max(1, c.start ?? 1);
    const restart = !started || (c.restart === "newPage" && l.page !== lastPage) || (c.restart === "newSection" && l.section !== lastSection) || (c.restart === "continuous" && l.section !== lastSection && !cfg(lastSection));
    if (restart) n = start - 1;
    started = true;
    lastSection = l.section;
    lastPage = l.page;
    if (l.suppress) continue;
    n++;
    const by = Math.max(1, c.countBy ?? 1);
    if (n % by === 0) out.push({ line: l, n });
  }
  return out;
}

/** Bir metin bloğundaki metin parçalarının ekran kutuları (sayfalama parçacıkları hariç). */
export function textRects(el: HTMLElement): DOMRect[] {
  const rects: DOMRect[] = [];
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, {
    acceptNode: (n) => {
      const p = n.parentElement;
      if (!p || p.closest(".bl-pg, [contenteditable='false']") || !(n.textContent ?? "").length) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    },
  });
  const range = document.createRange();
  for (let t = walker.nextNode(); t; t = walker.nextNode()) {
    range.selectNodeContents(t);
    for (const r of Array.from(range.getClientRects())) if (r.height > 0) rects.push(r);
  }
  return rects;
}

/** Kutuları satırlara gruplar: aynı sütunda dikey merkezi yakın olanlar tek satırdır. */
export function groupLines<T extends { top: number; bottom: number; left: number }>(rects: T[], colAt: (r: T) => number): { top: number; bottom: number; left: number; col: number }[] {
  const lines: { top: number; bottom: number; left: number; col: number }[] = [];
  for (const r of [...rects].sort((a, b) => a.top - b.top || a.left - b.left)) {
    const c = (r.top + r.bottom) / 2;
    const col = colAt(r);
    const hit = lines.find((l) => l.col === col && Math.abs((l.top + l.bottom) / 2 - c) < Math.max(2, (r.bottom - r.top) * 0.45));
    if (hit) {
      hit.top = Math.min(hit.top, r.top);
      hit.bottom = Math.max(hit.bottom, r.bottom);
      hit.left = Math.min(hit.left, r.left);
    } else lines.push({ top: r.top, bottom: r.bottom, left: r.left, col });
  }
  return lines.sort((a, b) => a.col - b.col || a.top - b.top);
}

const SKIP = new Set(["table", "textBox", "toc", "bibliography", "indexBlock"]);

/** Numaralanacak bölümlerdeki satır kutuları (belge sırasıyla; çok sütunluda sütun sütun). */
export function measureLines(
  view: EditorView,
  host: HTMLElement,
  scale: number,
  numbered: (section: number) => boolean,
  pageAt: (x: number, y: number) => number,
  colAt: (page: number, x: number) => number,
): (LineBox & { col: number })[] {
  const hr = host.getBoundingClientRect();
  const out: (LineBox & { col: number })[] = [];
  view.state.doc.forEach((sec, off, si) => {
    if (!numbered(si)) return;
    const secLines: (LineBox & { col: number })[] = [];
    sec.descendants((n, p) => {
      if (SKIP.has(n.type.name)) return false;
      if (!n.isTextblock) return true;
      const el = view.nodeDOM(off + 1 + p);
      if (!(el instanceof HTMLElement) || el.offsetParent === null) return false;
      const local = textRects(el).map((r) => ({ top: (r.top - hr.top) / scale, bottom: (r.bottom - hr.top) / scale, left: (r.left - hr.left) / scale }));
      if (!local.length) {
        // Boş paragraf da bir satırdır.
        const b = el.getBoundingClientRect();
        const lh = parseFloat(getComputedStyle(el).lineHeight) || b.height / scale;
        local.push({ top: (b.top - hr.top) / scale, bottom: (b.top - hr.top) / scale + Math.min(lh, b.height / scale), left: (b.left - hr.left) / scale });
      }
      const suppress = !!n.attrs.suppressLineNumbers;
      for (const l of groupLines(local, (r) => {
        const pg = pageAt(r.left, (r.top + r.bottom) / 2);
        return pg * 100 + colAt(pg, r.left);
      })) {
        const page = Math.floor(l.col / 100);
        secLines.push({ top: l.top, height: l.bottom - l.top, left: l.left, section: si, page, suppress, col: l.col });
      }
      return false;
    });
    // Sayfa ve sütun sırası (çok sütunlu bölümde önce birinci sütunun tüm satırları).
    secLines.sort((a, b) => a.col - b.col || a.top - b.top);
    out.push(...secLines);
  });
  return out;
}
