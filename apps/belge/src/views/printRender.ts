// Yazdırma görünümü: her sayfa, ekrandaki düzenin o sayfaya düşen kısmının kırpılmış bir kopyasıdır.
// Böylece tarayıcının kendi sayfa bölme davranışından bağımsız olarak ekranla birebir çıktı alınır.
// Farklı boyuttaki sayfalar (yatay bölümler) adlandırılmış @page kurallarıyla basılır.
import type { PageInfo } from "../editor/pagination";

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

/** Canlı sayfa kapsayıcısından (".bl-pages") yazdırma kökünü kurar. */
export function buildPrintRoot(pagesEl: HTMLElement, g: PrintGeometry): HTMLElement | null {
  removePrintRoot();
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

  // Sayfa boyutları için adlandırılmış @page kuralları.
  const style = document.createElement("style");
  style.id = STYLE_ID;
  const seen = new Set<string>();
  let css = "";
  for (const p of g.pages) {
    const name = pageName(p);
    if (seen.has(name)) continue;
    seen.add(name);
    css += `@page ${name} { size: ${p.w}px ${p.h}px; margin: 0; }\n`;
  }
  css += `@page { size: ${g.pages[0].w}px ${g.pages[0].h}px; margin: 0; }\n`;
  style.textContent = css;
  document.head.appendChild(style);

  const root = document.createElement("div");
  root.id = ROOT_ID;
  root.className = "bl-print-root";
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
    root.appendChild(page);
  });
  // Seçim ve düzenleme işaretlerini temizle.
  root.querySelectorAll(".ProseMirror-selectednode").forEach((e) => e.classList.remove("ProseMirror-selectednode"));
  root.querySelectorAll("[data-resize-handle], .bl-tb-handles").forEach((e) => e.remove());
  root.querySelectorAll("[contenteditable]").forEach((e) => e.removeAttribute("contenteditable"));
  document.body.appendChild(root);
  return root;
}
