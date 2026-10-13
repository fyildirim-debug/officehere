// RTF resimleri ({\pict}): \pngblip / \jpegblip onaltılık ya da \bin verisi → resim düğümü.
import type { JSONContent } from "@tiptap/core";
import { imageNode } from "../common";

const KINDS = new Set(["pngblip", "jpegblip", "emfblip", "wmetafile", "dibitmap", "wbitmap"]);

export class RtfPict {
  kind = "";
  hex: string[] = [];
  hexLen = 0;
  bin: Uint8Array | null = null;
  wg = 0;
  hg = 0;
  sx = 100;
  sy = 100;

  word(w: string, param: number | null) {
    if (KINDS.has(w)) this.kind = w;
    else if (w === "picwgoal") this.wg = param ?? 0;
    else if (w === "pichgoal") this.hg = param ?? 0;
    else if (w === "picscalex") this.sx = param ?? 100;
    else if (w === "picscaley") this.sy = param ?? 100;
  }

  text(s: string) {
    // 64 MB sınırı (bozuk dosyada bellek taşmasın).
    if (this.hexLen > 128 * 1024 * 1024) return;
    const h = s.replace(/[^0-9a-fA-F]/g, "");
    this.hex.push(h);
    this.hexLen += h.length;
  }

  /** Desteklenen biçimde (PNG/JPEG) resim düğümü; diğerleri atlanır. */
  node(fallbackW = 0, fallbackH = 0): JSONContent | null {
    if (this.kind !== "pngblip" && this.kind !== "jpegblip") return null;
    let bytes = this.bin;
    if (!bytes) {
      const hex = this.hex.join("");
      bytes = new Uint8Array(Math.floor(hex.length / 2));
      for (let i = 0; i < bytes.length; i++) bytes[i] = parseInt(hex.substr(i * 2, 2), 16);
    }
    const w = this.wg ? (this.wg * this.sx) / 100 / 15 : fallbackW > 0 ? fallbackW : 0;
    const h = this.hg ? (this.hg * this.sy) / 100 / 15 : fallbackH > 0 ? fallbackH : 0;
    return imageNode(bytes, { width: w || null, height: h || null });
  }
}
