/** Resim dosyası işlemleri: dosya seçtirme, Resmi Değiştir, Resim Olarak Kaydet. */
import { saveFileAs } from "@officehere/core";
import type { Controller } from "../controller";

/** Resim dosyası seçtirir; okunan resim (dosya adı, data: adresi, doğal boyut) geri çağrıya verilir. */
export function pickImage(ctl: Controller, done: (name: string, src: string, w: number, h: number) => void): void {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = "image/png,image/jpeg,image/gif,image/svg+xml,image/webp,image/bmp";
  input.onchange = () => {
    const f = input.files?.[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = () => {
      const src = String(reader.result);
      const im = new Image();
      im.onload = () => done(f.name, src, im.naturalWidth || 200, im.naturalHeight || 150);
      im.onerror = () => ctl.alert("Resim dosyası okunamadı.");
      im.src = src;
    };
    reader.readAsDataURL(f);
  };
  input.click();
}

/** Resmi Değiştir: yeni resim eskisinin kutusuna oranı korunarak sığdırılır; konum ve ad korunur. */
export function replaceImage(ctl: Controller, id: string): void {
  const sheet = ctl.sheet;
  pickImage(ctl, (_name, src, w, h) => {
    if (!sheet.images.some((x) => x.id === id)) return;
    ctl.ops.sheetOp(sheet, "Resmi Değiştir", () => {
      const img = sheet.images.find((x) => x.id === id);
      if (!img) return;
      const sc = Math.min(img.width / w, img.height / h);
      img.src = src;
      img.width = Math.max(20, Math.round(w * sc));
      img.height = Math.max(20, Math.round(h * sc));
    });
  });
}

const EXT: Record<string, string> = {
  "image/png": "png",
  "image/jpeg": "jpg",
  "image/gif": "gif",
  "image/svg+xml": "svg",
  "image/webp": "webp",
  "image/bmp": "bmp",
  "image/tiff": "tif",
  "image/x-emf": "emf",
  "image/x-wmf": "wmf",
};

/** data: adresini Blob'a çevirir (diğer adresler için fetch). */
export async function srcToBlob(src: string): Promise<Blob> {
  const m = /^data:([^;,]*)(;base64)?,(.*)$/s.exec(src);
  if (!m) return (await fetch(src)).blob();
  const type = m[1] || "application/octet-stream";
  if (!m[2]) return new Blob([decodeURIComponent(m[3])], { type });
  const bin = atob(m[3]);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new Blob([bytes], { type });
}

/** Kaydedilecek dosyanın adı: resmin adından (varsa eski uzantı atılır) ve türünden. */
export function imageFileName(name: string, type: string): string {
  const ext = EXT[type] ?? "png";
  const base = name.replace(/\.(png|jpe?g|gif|svg|webp|bmp|tiff?|emf|wmf)$/i, "").replace(/[\\/:*?"<>|]+/g, " ").trim() || "Resim";
  return `${base}.${ext}`;
}

/** Resim Olarak Kaydet: resmin özgün verisi dosyaya yazılır. */
export async function saveImageAs(ctl: Controller, id: string): Promise<void> {
  const img = ctl.sheet.images.find((x) => x.id === id);
  if (!img?.src) return;
  try {
    const blob = await srcToBlob(img.src);
    const name = imageFileName(img.name, blob.type);
    await saveFileAs(name, [{ name: "Resim", extensions: [name.slice(name.lastIndexOf(".") + 1)] }], blob);
  } catch {
    ctl.alert("Resim kaydedilemedi.");
  }
}
