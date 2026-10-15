/** Slayt resmi (PNG) dışa aktarma, ana hat metni ve masaüstünde doğrudan PDF. */
import { createRoot } from "react-dom/client";
import { downloadBlob, getNative, type PdfExportOptions } from "@officehere/core";
import { loadIo } from "../io/formats";
import type { Presentation, Shape } from "../model/types";
import { SlideView } from "../render/SlideView";
import slideCss from "../styles/slide.css?raw";

async function waitImages(el: HTMLElement) {
  const imgs = Array.from(el.querySelectorAll("img"));
  await Promise.all(imgs.map((i) => (i.complete ? Promise.resolve() : new Promise((r) => ((i.onload = r), (i.onerror = r))))));
}

/** Slaydı PNG'ye çizer (SVG foreignObject üzerinden). */
export async function slideToPng(pres: Presentation, index: number, scale = 1.5): Promise<Blob | null> {
  const canvas = await slideToCanvas(pres, index, scale);
  return new Promise((resolve) => canvas.toBlob((b) => resolve(b), "image/png"));
}

/**
 * Slaydı tuvale çizer (SVG foreignObject üzerinden). shapeClass ile şekillere sınıf
 * verilebilir (ör. henüz görünmeyen animasyonlu şekilleri gizlemek için).
 */
export async function slideToCanvas(pres: Presentation, index: number, scale = 1.5, shapeClass?: (s: Shape) => string | undefined): Promise<HTMLCanvasElement> {
  const host = document.createElement("div");
  host.style.cssText = `position:fixed;left:-100000px;top:0;width:${pres.width}px;height:${pres.height}px;`;
  document.body.appendChild(host);
  const root = createRoot(host);
  root.render(<SlideView pres={pres} slide={pres.slides[index]} slideNumber={index + 1} mode="print" shapeClass={shapeClass} />);
  await new Promise((r) => setTimeout(r, 60));
  await waitImages(host);
  const html = new XMLSerializer().serializeToString(host.firstElementChild!);
  root.unmount();
  host.remove();
  const W = pres.width;
  const H = pres.height;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W * scale}" height="${H * scale}" viewBox="0 0 ${W} ${H}"><foreignObject x="0" y="0" width="${W}" height="${H}"><div xmlns="http://www.w3.org/1999/xhtml" style="width:${W}px;height:${H}px"><style>${slideCss}</style>${html}</div></foreignObject></svg>`;
  const img = new Image();
  img.decoding = "sync";
  const url = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  await new Promise<void>((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error("Slayt resmi oluşturulamadı"));
    img.src = url;
  });
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(W * scale);
  canvas.height = Math.round(H * scale);
  const ctx = canvas.getContext("2d")!;
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
  return canvas;
}

export async function exportPngs(pres: Presentation, baseName: string, which: "current" | "all", current: number) {
  if (which === "current") {
    const b = await slideToPng(pres, current);
    if (b) downloadBlob(b, `${baseName}-Slayt${current + 1}.png`);
    return;
  }
  // JSZip yalnız çok slaytlı dışa aktarmada gerekir: ayrı parçadan yüklenir.
  const { JSZip } = await loadIo();
  const zip = new JSZip();
  for (let i = 0; i < pres.slides.length; i++) {
    const b = await slideToPng(pres, i);
    if (b) zip.file(`Slayt${i + 1}.png`, b);
  }
  downloadBlob(await zip.generateAsync({ type: "blob" }), `${baseName}-Slaytlar.zip`);
}

/** Ana hat: slayt başlıkları ve metinleri düz metin olarak. */
export function outlineText(pres: Presentation): string {
  const lines: string[] = [];
  pres.slides.forEach((s, i) => {
    lines.push(`Slayt ${i + 1}`);
    for (const sh of s.shapes) {
      if (sh.kind !== "sp" || !sh.text) continue;
      for (const p of sh.text.paragraphs) {
        const t = p.runs.map((r) => r.text.replace(/\v/g, " ")).join("").trim();
        if (t) lines.push(`${"\t".repeat(p.level + 1)}${t}`);
      }
    }
    if (s.notes.trim()) lines.push(`\tNotlar: ${s.notes.replace(/\n/g, " ")}`);
    lines.push("");
  });
  return lines.join("\r\n");
}

/** Masaüstü kabuğunda yazdırma sayfalarını doğrudan PDF'e yazar (size: sayfa boyutu, 96 dpi px; options: PDF seçenekleri). */
export async function exportPdfNative(size: { w: number; h: number }, baseName: string, options?: PdfExportOptions): Promise<string | null> {
  const native = getNative();
  if (!native?.exportPdf) return null;
  const w = size.w / 96;
  const h = size.h / 96;
  // Kâğıt boyutu dikey verilir, yönlendirme "landscape" ile döndürülür; @page boyutu önceliklidir.
  const r = await native.exportPdf(`${baseName}.pdf`, { ...options, landscape: w >= h, pageSize: { width: Math.min(w, h), height: Math.max(w, h) } });
  return r?.name ?? null;
}
