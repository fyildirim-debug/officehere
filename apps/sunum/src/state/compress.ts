/**
 * Resimleri Sıkıştır: resimler slayttaki görüntülenme boyutuna ve seçilen
 * çözünürlüğe (ppi) göre tuvalde yeniden örneklenir; istenirse kırpılan
 * alanlar silinir. Saydamlık içeren biçimler PNG, diğerleri JPEG kalır.
 */
import { flatShapes } from "../model/ops";
import type { PicShape, Shape } from "../model/types";
import type { Actions } from "./actions";

export interface CompressOptions {
  onlySelected: boolean;
  deleteCrop: boolean;
  /** 0: yüksek kalite (özgün çözünürlük korunur, yalnız kırpma silinir). */
  ppi: number;
}

export const COMPRESS_PPI: { ppi: number; label: string }[] = [
  { ppi: 0, label: "Yüksek kalite: Özgün resmin kalitesini korur" },
  { ppi: 330, label: "HD (330 ppi): Yüksek çözünürlüklü ekranlar için" },
  { ppi: 220, label: "Yazdır (220 ppi): Çoğu yazıcı ve ekranda mükemmel kalite" },
  { ppi: 150, label: "Web (150 ppi): Web sayfaları ve projektörler için iyi" },
  { ppi: 96, label: "E-posta (96 ppi): Paylaşılacak belge boyutunu en aza indirir" },
];

/** Hedef piksel boyutu: görüntülenme boyutu (px, 96 dpi) × ppi / 96; özgünden büyük olmaz. */
export function targetSize(natW: number, natH: number, shownW: number, shownH: number, crop: PicShape["crop"], ppi: number, deleteCrop: boolean): { sx: number; sy: number; sw: number; sh: number; w: number; h: number } {
  const c = deleteCrop && crop ? crop : { l: 0, t: 0, r: 0, b: 0 };
  const sx = Math.round(natW * c.l);
  const sy = Math.round(natH * c.t);
  const sw = Math.max(1, Math.round(natW * (1 - c.l - c.r)));
  const sh = Math.max(1, Math.round(natH * (1 - c.t - c.b)));
  // Kırpma silinmiyorsa resmin tamamı, görünen bölümün oranında büyütülmüş boyutta tutulur.
  const fullW = deleteCrop ? shownW : shownW / Math.max(0.01, 1 - (crop?.l ?? 0) - (crop?.r ?? 0));
  const fullH = deleteCrop ? shownH : shownH / Math.max(0.01, 1 - (crop?.t ?? 0) - (crop?.b ?? 0));
  const k = ppi ? Math.min(1, ((fullW * ppi) / 96) / sw, ((fullH * ppi) / 96) / sh) : 1;
  return { sx, sy, sw, sh, w: Math.max(1, Math.round(sw * k)), h: Math.max(1, Math.round(sh * k)) };
}

const loadImage = (src: string) =>
  new Promise<HTMLImageElement>((res, rej) => {
    const img = new Image();
    img.onload = () => res(img);
    img.onerror = () => rej(new Error("Resim yüklenemedi"));
    img.src = src;
  });

async function compressOne(p: PicShape, o: CompressOptions): Promise<{ src: string; cropped: boolean } | null> {
  if (p.media || !p.src.startsWith("data:image/") || p.src.startsWith("data:image/svg")) return null;
  const img = await loadImage(p.src);
  const t = targetSize(img.naturalWidth, img.naturalHeight, p.w, p.h, p.crop, o.ppi, o.deleteCrop);
  const cropping = o.deleteCrop && !!p.crop && (p.crop.l > 0 || p.crop.t > 0 || p.crop.r > 0 || p.crop.b > 0);
  if (!cropping && t.w >= img.naturalWidth && t.h >= img.naturalHeight) return null;
  const canvas = document.createElement("canvas");
  canvas.width = t.w;
  canvas.height = t.h;
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;
  ctx.imageSmoothingQuality = "high";
  ctx.drawImage(img, t.sx, t.sy, t.sw, t.sh, 0, 0, t.w, t.h);
  const jpeg = /^data:image\/jpe?g/i.test(p.src);
  const src = jpeg ? canvas.toDataURL("image/jpeg", 0.85) : canvas.toDataURL("image/png");
  // Yeniden kodlama dosyayı büyütürse (ve kırpma yoksa) özgün korunur.
  if (!cropping && src.length >= p.src.length) return null;
  return { src, cropped: cropping };
}

export async function compressPictures(a: Actions, o: CompressOptions): Promise<{ count: number; before: number; after: number }> {
  const sel = new Set(a.s.sel);
  const targets: PicShape[] = [];
  a.pres.slides.forEach((sl, i) => {
    if (o.onlySelected && i !== a.s.slideIdx) return;
    for (const s of flatShapes(sl.shapes)) if (s.kind === "pic" && (!o.onlySelected || sel.has(s.id))) targets.push(s);
  });
  const results = new Map<string, { src: string; cropped: boolean }>();
  let before = 0;
  let after = 0;
  for (const p of targets) {
    try {
      const r = await compressOne(p, o);
      if (!r) continue;
      results.set(p.id, r);
      before += p.src.length * 0.75;
      after += r.src.length * 0.75;
    } catch {
      // okunamayan resim atlanır
    }
  }
  if (results.size) {
    const map = (s: Shape): Shape => {
      if (s.kind === "group") return { ...s, children: s.children.map(map) };
      const r = s.kind === "pic" ? results.get(s.id) : undefined;
      return r && s.kind === "pic" ? { ...s, src: r.src, crop: r.cropped ? undefined : s.crop } : s;
    };
    a.store.commit((p) => ({ ...p, slides: p.slides.map((sl) => ({ ...sl, shapes: sl.shapes.map(map) })) }));
  }
  return { count: results.size, before, after };
}
