// Yazdırma görünümü: her sayfa, ekrandaki düzenin o sayfaya düşen kısmının kırpılmış bir kopyasıdır.
// Böylece tarayıcının kendi sayfa bölme davranışından bağımsız olarak ekranla birebir çıktı alınır.
// Farklı boyuttaki sayfalar (yatay bölümler) adlandırılmış @page kurallarıyla basılır.
import type { PageInfo } from "../editor/pagination";
import { sheetGrid } from "../print/printOptions";

export interface PrintGeometry {
  pages: PageInfo[];
  /** Ekran ölçeği (yakınlaştırma). */
  scale: number;
  pageColor: string | null;
}

const ROOT_ID = "bl-print-root";
const STYLE_ID = "bl-page-style";

export function removePrintRoot() {
  document.getElementById(ROOT_ID)?.remove();
  document.getElementById(STYLE_ID)?.remove();
}

const pageName = (p: PageInfo) => `bl-${Math.round(p.w)}x${Math.round(p.h)}`;

interface FlowItem {
  el: HTMLElement;
  /** Bölüm kabuğu (tek sütunlu bölümlerin çocukları için). */
  section: HTMLElement | null;
  top: number;
  bottom: number;
}

/**
 * Canlı sayfa kapsayıcısından (".bl-pages") her sayfanın bağımsız kopyasını kurar (belgeye eklenmez).
 * Yazdırma kökü ve Dosya › Yazdır önizlemesi bu kopyaları kullanır.
 */
export function renderPages(pagesEl: HTMLElement, g: PrintGeometry): HTMLElement[] | null {
  const body = pagesEl.querySelector<HTMLElement>(".bl-body");
  const pm = body?.querySelector<HTMLElement>(".ProseMirror");
  if (!body || !pm || !g.pages.length) return null;
  const base = pagesEl.getBoundingClientRect();
  const Y = (screenY: number) => (screenY - base.top) / g.scale;
  const pmRect = pm.getBoundingClientRect();
  const contentTop = Y(pmRect.top) + (parseFloat(getComputedStyle(pm).paddingTop) || 0);

  // Akış öğeleri: tek sütunlu bölümlerin blokları (ve ara parçaları), çok sütunlu bölümün kendisi.
  const items: FlowItem[] = [];
  const box = (el: HTMLElement, section: HTMLElement | null) => {
    const r = el.getBoundingClientRect();
    const cs = getComputedStyle(el);
    items.push({ el, section, top: Y(r.top) - (parseFloat(cs.marginTop) || 0), bottom: Y(r.bottom) + (parseFloat(cs.marginBottom) || 0) });
  };
  for (const child of Array.from(pm.children) as HTMLElement[]) {
    if (child.classList.contains("bl-section") && !child.classList.contains("is-cols")) {
      for (const k of Array.from(child.children) as HTMLElement[]) box(k, child);
    } else box(child, null);
  }

  const out: HTMLElement[] = [];
  g.pages.forEach((pg, i) => {
    const pTop = pg.top;
    const pBottom = pTop + pg.h;
    const page = document.createElement("div");
    page.className = "bl-print-page";
    page.style.width = `${pg.w}px`;
    page.style.height = `${pg.h}px`;
    (page.style as CSSStyleDeclaration & { page?: string }).page = pageName(pg);
    page.style.setProperty("page", pageName(pg));
    if (g.pageColor) page.style.background = g.pageColor;
    // Sayfa zemini: üst/alt bilgi, filigran, kenarlık, dipnotlar.
    const sheet = pagesEl.querySelector<HTMLElement>(`.bl-page[data-page="${i}"]`);
    if (sheet) {
      const s = sheet.cloneNode(true) as HTMLElement;
      s.style.top = "0px";
      s.style.left = "0px";
      s.style.boxShadow = "none";
      page.appendChild(s);
    }
    // İçerik kopyası: yalnız bu sayfayla kesişen öğeler tutulur.
    const bodyClone = body.cloneNode(false) as HTMLElement;
    bodyClone.style.position = "absolute";
    bodyClone.style.top = `${contentTop - pTop}px`;
    bodyClone.style.left = `${-pg.left}px`;
    bodyClone.style.paddingTop = "0px";
    bodyClone.style.height = "auto";
    const pmClone = pm.cloneNode(false) as HTMLElement;
    pmClone.removeAttribute("contenteditable");
    pmClone.style.minHeight = "0px";
    let first = -1;
    let last = -1;
    items.forEach((b, k) => {
      if (b.bottom > pTop && b.top < pBottom) {
        if (first < 0) first = k;
        last = k;
      }
    });
    if (first >= 0) {
      const spacer = document.createElement("div");
      spacer.style.height = `${Math.max(0, items[first].top - contentTop)}px`;
      spacer.style.flex = "none";
      pmClone.appendChild(spacer);
      let shell: HTMLElement | null = null;
      let shellOf: HTMLElement | null = null;
      for (let k = first; k <= last; k++) {
        const it = items[k];
        if (it.section) {
          if (shellOf !== it.section) {
            shell = it.section.cloneNode(false) as HTMLElement;
            shellOf = it.section;
            pmClone.appendChild(shell);
          }
          shell!.appendChild(it.el.cloneNode(true));
        } else {
          shell = null;
          shellOf = null;
          pmClone.appendChild(it.el.cloneNode(true));
        }
      }
    }
    bodyClone.appendChild(pmClone);
    page.appendChild(bodyClone);
    // Seçim ve düzenleme işaretlerini temizle.
    page.querySelectorAll(".ProseMirror-selectednode").forEach((e) => e.classList.remove("ProseMirror-selectednode"));
    page.querySelectorAll("[data-resize-handle], .bl-tb-handles, .bl-tsp, .bl-pg-tilerow").forEach((e) => e.remove());
    page.querySelectorAll("[contenteditable]").forEach((e) => e.removeAttribute("contenteditable"));
    page.querySelectorAll("[id]").forEach((e) => e.removeAttribute("id"));
    out.push(page);
  });
  return out;
}

/** Yaprak: birden çok sayfa küçültülerek ızgaraya dizilir (yaprak başına sayfa). */
export function buildSheet(pages: HTMLElement[], infos: PageInfo[], perSheet: number): HTMLElement {
  const first = infos[0];
  const base = sheetGrid(perSheet);
  const pad = 12;
  // Kağıt (ilk sayfanın boyutu ya da döndürülmüşü) ve ızgara yönü, sayfalar en büyük görünecek biçimde seçilir.
  let best = { W: first.w, H: first.h, cols: base.cols, rows: base.rows, sc: 0 };
  for (const [W, H] of [
    [first.w, first.h],
    [first.h, first.w],
  ]) {
    for (const [cols, rows] of [
      [base.cols, base.rows],
      [base.rows, base.cols],
    ]) {
      const sc = Math.min((W - pad * (cols + 1)) / cols / first.w, (H - pad * (rows + 1)) / rows / first.h);
      if (sc > best.sc + 1e-6) best = { W, H, cols, rows, sc };
    }
  }
  const { W, H } = best;
  const grid = { cols: best.cols, rows: best.rows };
  const sheet = document.createElement("div");
  sheet.className = "bl-print-page bl-print-sheet";
  sheet.style.width = `${W}px`;
  sheet.style.height = `${H}px`;
  const name = `bl-${Math.round(W)}x${Math.round(H)}`;
  sheet.style.setProperty("page", name);
  const cw = (W - pad * (grid.cols + 1)) / grid.cols;
  const ch = (H - pad * (grid.rows + 1)) / grid.rows;
  pages.forEach((pg, k) => {
    const info = infos[k];
    const sc = Math.min(cw / info.w, ch / info.h);
    const col = k % grid.cols;
    const row = Math.floor(k / grid.cols);
    const cell = document.createElement("div");
    cell.className = "bl-print-cell";
    cell.style.left = `${pad + col * (cw + pad) + (cw - info.w * sc) / 2}px`;
    cell.style.top = `${pad + row * (ch + pad) + (ch - info.h * sc) / 2}px`;
    cell.style.width = `${info.w}px`;
    cell.style.height = `${info.h}px`;
    cell.style.transform = `scale(${sc})`;
    pg.style.removeProperty("page");
    cell.appendChild(pg);
    sheet.appendChild(cell);
  });
  return sheet;
}

/** Yapraklar (sayfa sıraları) için yazdırma sayfa öğeleri. */
export function assembleSheets(pages: HTMLElement[], infos: PageInfo[], sheets: number[][], perSheet: number): HTMLElement[] {
  return sheets.map((list) => {
    if (perSheet <= 1) return pages[list[0]].cloneNode(true) as HTMLElement;
    return buildSheet(
      list.map((i) => pages[i].cloneNode(true) as HTMLElement),
      list.map((i) => infos[i]),
      perSheet,
    );
  });
}

/** Canlı sayfa kapsayıcısından yazdırma kökünü kurar; `sheets` verilirse yalnız o yapraklar basılır. */
export function buildPrintRoot(pagesEl: HTMLElement, g: PrintGeometry, layout?: { sheets: number[][]; perSheet: number }): HTMLElement | null {
  removePrintRoot();
  const pages = renderPages(pagesEl, g);
  if (!pages) return null;
  const sheets = layout?.sheets ?? pages.map((_, i) => [i]);
  const items = assembleSheets(pages, g.pages, sheets, layout?.perSheet ?? 1);
  if (!items.length) return null;
  // Sayfa boyutları için adlandırılmış @page kuralları.
  const style = document.createElement("style");
  style.id = STYLE_ID;
  const sizes = new Map<string, string>();
  for (const el of items) {
    const name = el.style.getPropertyValue("page") || `bl-${Math.round(parseFloat(el.style.width))}x${Math.round(parseFloat(el.style.height))}`;
    el.style.setProperty("page", name);
    sizes.set(name, `${el.style.width} ${el.style.height}`);
  }
  let css = "";
  for (const [name, size] of sizes) css += `@page ${name} { size: ${size}; margin: 0; }\n`;
  css += `@page { size: ${items[0].style.width} ${items[0].style.height}; margin: 0; }\n`;
  style.textContent = css;
  document.head.appendChild(style);
  const root = document.createElement("div");
  root.id = ROOT_ID;
  root.className = "bl-print-root";
  for (const el of items) root.appendChild(el);
  document.body.appendChild(root);
  return root;
}
