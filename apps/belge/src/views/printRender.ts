// Yazdırma görünümü: her sayfa, ekrandaki düzenin o sayfaya düşen kısmının kırpılmış bir kopyasıdır.
// Böylece tarayıcının kendi sayfa bölme davranışından bağımsız olarak ekranla birebir çıktı alınır.

export interface PrintGeometry {
  pageCount: number;
  /** Sayfa yüksekliği / genişliği (px, ölçeksiz). */
  pageH: number;
  pageW: number;
  /** Sayfa başlangıçları arası mesafe (sayfa + boşluk). */
  stride: number;
  /** Ekran ölçeği (yakınlaştırma). */
  scale: number;
  columns: number;
  pageColor: string | null;
}

const ROOT_ID = "bl-print-root";

export function removePrintRoot() {
  document.getElementById(ROOT_ID)?.remove();
}

/** Canlı sayfa kapsayıcısından (".bl-pages") yazdırma kökünü kurar. */
export function buildPrintRoot(pagesEl: HTMLElement, g: PrintGeometry): HTMLElement | null {
  removePrintRoot();
  const body = pagesEl.querySelector<HTMLElement>(".bl-body");
  const pm = body?.querySelector<HTMLElement>(".ProseMirror");
  if (!body || !pm) return null;
  const base = pagesEl.getBoundingClientRect();
  const Y = (screenY: number) => (screenY - base.top) / g.scale;
  const pmRect = pm.getBoundingClientRect();
  const contentTop = Y(pmRect.top) + (parseFloat(getComputedStyle(pm).paddingTop) || 0);

  // Üst düzey blokların konumları (görsel; sütun düzeninde dönüşümler dahil).
  const kids = Array.from(pm.children) as HTMLElement[];
  const boxes = kids.map((el) => {
    const r = el.getBoundingClientRect();
    const cs = getComputedStyle(el);
    return { top: Y(r.top) - (parseFloat(cs.marginTop) || 0), bottom: Y(r.bottom) + (parseFloat(cs.marginBottom) || 0) };
  });

  const root = document.createElement("div");
  root.id = ROOT_ID;
  root.className = "bl-print-root";
  for (let i = 0; i < g.pageCount; i++) {
    const pTop = i * g.stride;
    const pBottom = pTop + g.pageH;
    const page = document.createElement("div");
    page.className = "bl-print-page";
    page.style.width = `${g.pageW}px`;
    page.style.height = `${g.pageH}px`;
    if (g.pageColor) page.style.background = g.pageColor;
    // Sayfa zemini: üst/alt bilgi, filigran, kenarlık.
    const sheet = pagesEl.querySelector<HTMLElement>(`.bl-page[data-page="${i}"]`);
    if (sheet) {
      const s = sheet.cloneNode(true) as HTMLElement;
      s.style.top = "0px";
      s.style.boxShadow = "none";
      page.appendChild(s);
    }
    // İçerik kopyası: yalnız bu sayfayla kesişen bloklar tutulur.
    const bodyClone = body.cloneNode(false) as HTMLElement;
    bodyClone.style.position = "absolute";
    bodyClone.style.top = `${contentTop - pTop}px`;
    bodyClone.style.paddingTop = "0px";
    bodyClone.style.height = "auto";
    const pmClone = pm.cloneNode(false) as HTMLElement;
    pmClone.removeAttribute("contenteditable");
    pmClone.style.minHeight = "0px";
    let first = -1;
    let last = -1;
    boxes.forEach((b, k) => {
      const hit = g.columns > 1 ? b.bottom > pTop - 1 && b.top < pBottom + 1 : b.bottom > pTop && b.top < pBottom;
      if (hit) {
        if (first < 0) first = k;
        last = k;
      }
    });
    if (g.columns > 1) {
      // Sütun düzeninde bloklar dönüşümle taşındığı için tüm akış korunur.
      kids.forEach((k) => pmClone.appendChild(k.cloneNode(true)));
    } else if (first >= 0) {
      const spacer = document.createElement("div");
      spacer.style.height = `${Math.max(0, boxes[first].top - contentTop)}px`;
      spacer.style.flex = "none";
      pmClone.appendChild(spacer);
      for (let k = first; k <= last; k++) pmClone.appendChild(kids[k].cloneNode(true));
    }
    bodyClone.appendChild(pmClone);
    page.appendChild(bodyClone);
    root.appendChild(page);
  }
  // Seçim ve düzenleme işaretlerini temizle.
  root.querySelectorAll(".ProseMirror-selectednode").forEach((e) => e.classList.remove("ProseMirror-selectednode"));
  root.querySelectorAll("[data-resize-handle]").forEach((e) => e.remove());
  document.body.appendChild(root);
  return root;
}
