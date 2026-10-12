// Satır içi resim: boyutlandırma tutamaçları, kenarlık, resim stili, döndürme ve alternatif metin.
import Image from "@tiptap/extension-image";

export const IMAGE_STYLES: { id: string; label: string }[] = [
  { id: "none", label: "Stil Yok" },
  { id: "frame", label: "Basit Çerçeve, Beyaz" },
  { id: "shadow", label: "Gölge Dikdörtgen" },
  { id: "rounded", label: "Yuvarlatılmış Çapraz Köşeli Dikdörtgen" },
  { id: "soft", label: "Yumuşak Kenarlı Dikdörtgen" },
  { id: "oval", label: "Oval" },
  { id: "thick", label: "Kalın Mat, Siyah" },
];

export const BelgeImage = Image.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      border: {
        default: null,
        parseHTML: (el) => el.getAttribute("data-border"),
        renderHTML: (a) => (a.border ? { "data-border": a.border } : {}),
      },
      imgStyle: {
        default: null,
        parseHTML: (el) => el.getAttribute("data-img-style"),
        renderHTML: (a) => (a.imgStyle && a.imgStyle !== "none" ? { "data-img-style": a.imgStyle } : {}),
      },
      rotate: {
        default: 0,
        parseHTML: (el) => Number(el.getAttribute("data-rotate") ?? 0),
        renderHTML: (a) => (a.rotate ? { "data-rotate": String(a.rotate) } : {}),
      },
      flipH: { default: false, parseHTML: (el) => el.hasAttribute("data-flip-h"), renderHTML: (a) => (a.flipH ? { "data-flip-h": "" } : {}) },
      flipV: {
        default: false,
        parseHTML: (el) => el.hasAttribute("data-flip-v"),
        renderHTML: (a) => {
          const t: string[] = [];
          if (a.rotate) t.push(`rotate(${a.rotate}deg)`);
          if (a.flipH) t.push("scaleX(-1)");
          if (a.flipV) t.push("scaleY(-1)");
          const out: Record<string, string> = {};
          if (a.flipV) out["data-flip-v"] = "";
          const style: string[] = [];
          if (t.length) style.push(`transform:${t.join(" ")}`);
          if (a.border) style.push(`outline:${a.border}`);
          if (style.length) out.style = style.join(";");
          return out;
        },
      },
    };
  },
}).configure({
  inline: true,
  allowBase64: true,
  resize: { enabled: true, directions: ["top-left", "top-right", "bottom-left", "bottom-right"], minWidth: 16, minHeight: 16, alwaysPreserveAspectRatio: true },
});

/** Dosyadan resim okuyup doğal boyutunu (en fazla verilen genişlikte) döner. */
export function readImageFile(file: File, maxWidth: number): Promise<{ src: string; width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => {
      const src = String(fr.result);
      const img = new window.Image();
      img.onload = () => {
        let w = img.naturalWidth || 200;
        let h = img.naturalHeight || 150;
        if (w > maxWidth) {
          h = Math.round((h * maxWidth) / w);
          w = Math.round(maxWidth);
        }
        resolve({ src, width: w, height: h });
      };
      img.onerror = () => reject(new Error("Resim okunamadı"));
      img.src = src;
    };
    fr.onerror = () => reject(fr.error);
    fr.readAsDataURL(file);
  });
}

export type ShapeKind = "line" | "arrow" | "rect" | "roundRect" | "oval" | "triangle" | "star" | "rightArrow";

export const SHAPES: { id: ShapeKind; label: string }[] = [
  { id: "line", label: "Çizgi" },
  { id: "arrow", label: "Ok" },
  { id: "rect", label: "Dikdörtgen" },
  { id: "roundRect", label: "Yuvarlatılmış Dikdörtgen" },
  { id: "oval", label: "Oval" },
  { id: "triangle", label: "İkizkenar Üçgen" },
  { id: "rightArrow", label: "Sağ Ok" },
  { id: "star", label: "Yıldız: 5 Köşeli" },
];

/** Şekil SVG yolu (100×100 kutuda). */
export function shapeSvg(kind: ShapeKind, fill: string, stroke: string, w = 160, h = 100): string {
  const sw = 3;
  let body = "";
  switch (kind) {
    case "line":
      body = `<line x1="4" y1="${h - 4}" x2="${w - 4}" y2="4" stroke="${stroke}" stroke-width="${sw}"/>`;
      break;
    case "arrow":
      body = `<defs><marker id="a" markerWidth="10" markerHeight="10" refX="8" refY="5" orient="auto"><path d="M0,0 L10,5 L0,10 z" fill="${stroke}"/></marker></defs><line x1="4" y1="${h / 2}" x2="${w - 12}" y2="${h / 2}" stroke="${stroke}" stroke-width="${sw}" marker-end="url(#a)"/>`;
      break;
    case "rect":
      body = `<rect x="2" y="2" width="${w - 4}" height="${h - 4}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`;
      break;
    case "roundRect":
      body = `<rect x="2" y="2" width="${w - 4}" height="${h - 4}" rx="16" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`;
      break;
    case "oval":
      body = `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${w / 2 - 3}" ry="${h / 2 - 3}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`;
      break;
    case "triangle":
      body = `<polygon points="${w / 2},3 ${w - 3},${h - 3} 3,${h - 3}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`;
      break;
    case "rightArrow":
      body = `<polygon points="3,${h * 0.3} ${w * 0.62},${h * 0.3} ${w * 0.62},3 ${w - 3},${h / 2} ${w * 0.62},${h - 3} ${w * 0.62},${h * 0.7} 3,${h * 0.7}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`;
      break;
    case "star": {
      const pts: string[] = [];
      for (let i = 0; i < 10; i++) {
        const r = i % 2 === 0 ? Math.min(w, h) / 2 - 3 : Math.min(w, h) / 5;
        const a = (Math.PI / 5) * i - Math.PI / 2;
        pts.push(`${(w / 2 + r * Math.cos(a)).toFixed(1)},${(h / 2 + r * Math.sin(a)).toFixed(1)}`);
      }
      body = `<polygon points="${pts.join(" ")}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`;
      break;
    }
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">${body}</svg>`;
}

/** SVG'yi PNG veri URL'sine çevirir (DOCX'e gömülebilmesi için). */
export function svgToPng(svg: string, w: number, h: number): Promise<string> {
  return new Promise((resolve, reject) => {
    const img = new window.Image();
    img.onload = () => {
      const c = document.createElement("canvas");
      c.width = w * 2;
      c.height = h * 2;
      const ctx = c.getContext("2d");
      if (!ctx) return reject(new Error("canvas yok"));
      ctx.scale(2, 2);
      ctx.drawImage(img, 0, 0, w, h);
      resolve(c.toDataURL("image/png"));
    };
    img.onerror = () => reject(new Error("SVG çizilemedi"));
    img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  });
}
