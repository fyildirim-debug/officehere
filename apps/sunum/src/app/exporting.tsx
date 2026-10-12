/** Yazdırma/PDF (her slayt bir sayfa) ve PNG dışa aktarma. */
import { useEffect } from "react";
import { createPortal } from "react-dom";
import { createRoot } from "react-dom/client";
import JSZip from "jszip";
import { downloadBlob } from "@officehere/core";
import type { Presentation } from "../model/types";
import { SlideView } from "../render/SlideView";
import slideCss from "../styles/slide.css?raw";

/** Yazdırma kökü: yalnız yazdırmada görünür; @page boyutu slayt oranındadır. */
export function PrintRoot({ pres, onDone, includeHidden = false }: { pres: Presentation; onDone: () => void; includeHidden?: boolean }) {
  useEffect(() => {
    const style = document.createElement("style");
    style.textContent = `@page { size: ${(pres.width / 96) * 2.54}cm ${(pres.height / 96) * 2.54}cm; margin: 0; }`;
    document.head.appendChild(style);
    document.documentElement.classList.add("is-printing");
    const after = () => onDone();
    window.addEventListener("afterprint", after);
    const imgs = Array.from(document.querySelectorAll<HTMLImageElement>(".print-root img"));
    Promise.all(imgs.map((i) => (i.complete ? Promise.resolve() : new Promise((r) => ((i.onload = r), (i.onerror = r)))))).then(() =>
      requestAnimationFrame(() =>
        requestAnimationFrame(() => {
          window.print();
          // Bazı tarayıcılar afterprint göndermez.
          window.setTimeout(after, 500);
        }),
      ),
    );
    return () => {
      window.removeEventListener("afterprint", after);
      style.remove();
      document.documentElement.classList.remove("is-printing");
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  return createPortal(
    <div className="print-root">
      {pres.slides
        .map((s, i) => ({ s, i }))
        .filter(({ s }) => includeHidden || !s.hidden)
        .map(({ s, i }) => (
          <div key={s.id} className="print-page" style={{ width: pres.width, height: pres.height }}>
            <SlideView pres={pres} slide={s} slideNumber={i + 1} mode="print" />
          </div>
        ))}
    </div>,
    document.body,
  );
}

async function waitImages(el: HTMLElement) {
  const imgs = Array.from(el.querySelectorAll("img"));
  await Promise.all(imgs.map((i) => (i.complete ? Promise.resolve() : new Promise((r) => ((i.onload = r), (i.onerror = r))))));
}

/** Slaydı PNG'ye çizer (SVG foreignObject üzerinden). */
export async function slideToPng(pres: Presentation, index: number, scale = 1.5): Promise<Blob | null> {
  const host = document.createElement("div");
  host.style.cssText = `position:fixed;left:-100000px;top:0;width:${pres.width}px;height:${pres.height}px;`;
  document.body.appendChild(host);
  const root = createRoot(host);
  root.render(<SlideView pres={pres} slide={pres.slides[index]} slideNumber={index + 1} mode="print" />);
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
  return new Promise((resolve) => canvas.toBlob((b) => resolve(b), "image/png"));
}

export async function exportPngs(pres: Presentation, baseName: string, which: "current" | "all", current: number) {
  if (which === "current") {
    const b = await slideToPng(pres, current);
    if (b) downloadBlob(b, `${baseName}-Slayt${current + 1}.png`);
    return;
  }
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
