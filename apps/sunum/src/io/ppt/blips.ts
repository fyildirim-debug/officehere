/**
 * OfficeArt resim deposu (BStore): FBSE girdileri ve "Pictures" akışındaki
 * BLIP kayıtları. JPEG, PNG, DIB (BMP'ye çevrilir) ve TIFF okunur; meta
 * dosyalar (EMF/WMF/PICT) tarayıcıda gösterilemediği için atlanır.
 */
import { toBase64 } from "../odp/common";
import { RT, Stream, type Rec } from "./records";

const MAX_BLIP = 40 * 1024 * 1024;

/** BLIP kaydı → data: URL. */
function blipData(s: Stream, r: Rec | undefined): string | undefined {
  if (!r || r.len > MAX_BLIP) return undefined;
  const uids = 1 + (r.inst & 1);
  let mime: string;
  switch (r.type) {
    case 0xf01d:
    case 0xf02a:
      mime = "image/jpeg";
      break;
    case 0xf01e:
      mime = "image/png";
      break;
    case 0xf01f:
      mime = "image/bmp";
      break;
    case 0xf029:
      mime = "image/tiff";
      break;
    default:
      return undefined;
  }
  const off = r.start + 16 * uids + 1;
  if (off >= r.start + r.len) return undefined;
  let bytes = s.bytes.subarray(off, r.start + r.len);
  if (mime === "image/bmp") bytes = dibToBmp(bytes);
  return `data:${mime};base64,${toBase64(bytes)}`;
}

/** Aygıttan bağımsız bit eşlem (DIB) → BMP dosyası (14 baytlık başlık eklenir). */
export function dibToBmp(dib: Uint8Array): Uint8Array {
  const dv = new DataView(dib.buffer, dib.byteOffset, dib.byteLength);
  if (dib.length < 16) return dib;
  const hsize = dv.getUint32(0, true);
  const bits = hsize >= 16 ? dv.getUint16(14, true) : 24;
  const compression = hsize >= 20 ? dv.getUint32(16, true) : 0;
  const used = hsize >= 36 ? dv.getUint32(32, true) : 0;
  const colors = used || (bits <= 8 ? 1 << bits : 0);
  const masks = hsize === 40 && compression === 3 ? 12 : 0;
  const offset = 14 + hsize + masks + colors * 4;
  const out = new Uint8Array(14 + dib.length);
  const ov = new DataView(out.buffer);
  out[0] = 0x42;
  out[1] = 0x4d;
  ov.setUint32(2, out.length, true);
  ov.setUint32(10, Math.min(offset, out.length), true);
  out.set(dib, 14);
  return out;
}

/** Resim deposu: 1 tabanlı dizinle resim. */
export function blipStore(doc: Stream, dgg: Rec | undefined, pictures: Uint8Array | null): (pib: number) => string | undefined {
  const bstore = doc.kid(dgg, RT.BStoreContainer);
  const entries = doc.kids(bstore).filter((r) => r.type === RT.FBSE);
  const pics = pictures ? new Stream(pictures) : undefined;
  const cache = new Map<number, string | undefined>();
  return (pib) => {
    if (!pib || pib > entries.length) return undefined;
    if (cache.has(pib)) return cache.get(pib);
    const e = entries[pib - 1];
    let url: string | undefined;
    const size = doc.u32(e.start + 20);
    const delay = doc.u32(e.start + 28);
    const cbName = doc.u8(e.start + 33);
    const inner = e.start + 36 + cbName;
    if (e.len > 36 + cbName + 8) url = blipData(doc, doc.header(inner, e.start + e.len));
    else if (pics && size > 0 && delay < pics.length) url = blipData(pics, pics.header(delay));
    cache.set(pib, url);
    return url;
  };
}
