/** Resim komutları: Resmi Değiştir (dosyadan) ve Resim Olarak Kaydet. */
import { saveFileAs } from "@officehere/core";
import type { PicShape } from "../model/types";
import type { Actions } from "../state/actions";
import { readImageFile } from "./clipboard";

/** Seçili resimlerin yerine dosyadan seçilen resmi koyar: kutuya sığdırılır, kırpma sıfırlanır. */
export function replacePicture(a: Actions) {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = "image/*";
  input.onchange = async () => {
    const f = input.files?.[0];
    if (!f) return;
    try {
      const img = await readImageFile(f);
      a.updateShapes((s) => {
        if (s.kind !== "pic") return s;
        const sc = Math.min(s.w / img.w, s.h / img.h);
        return { ...s, src: img.src, crop: undefined, w: img.w * sc, h: img.h * sc };
      });
    } catch {
      a.message("Resim okunamadı.", "error");
    }
  };
  input.click();
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

/** Dosya adı için güvenli ad (yasak karakterler atılır). */
export function pictureFileName(name: string, type: string): string {
  const ext = EXT[type] ?? "png";
  const base = name.replace(/[\\/:*?"<>|]+/g, " ").trim() || "Resim";
  return `${base}.${ext}`;
}

/** Resmin özgün verisini dosyaya kaydeder (kırpma ve efektler uygulanmaz). */
export async function savePictureAs(a: Actions, pic: PicShape) {
  if (!pic.src) return;
  try {
    const blob = await srcToBlob(pic.src);
    const name = pictureFileName(pic.name, blob.type);
    const ext = name.slice(name.lastIndexOf(".") + 1);
    await saveFileAs(name, [{ name: "Resim", extensions: [ext] }], blob);
  } catch {
    a.message("Resim kaydedilemedi.", "error");
  }
}
