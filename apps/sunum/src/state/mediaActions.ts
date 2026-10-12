/** Grafik ve video/ses ekleme komutları (Ekle > Grafik, Ekle > Medya). */
import { makeChart } from "../model/chart";
import { uid } from "../model/factory";
import { updateSlideShapes } from "../model/ops";
import type { ChartData, ChartGrouping, ChartShape, ChartType, MediaInfo, PicShape } from "../model/types";
import { SPEAKER_POSTER } from "../render/MediaBody";
import type { Actions } from "./actions";

/** Grafik ekler (boş içerik yer tutucusu varsa onun yerine). Eklenen grafiğin kimliğini döndürür. */
export function insertChart(a: Actions, type: ChartType, grouping?: ChartGrouping): string | null {
  if (!a.slide()) a.newSlide();
  const p = a.pres;
  const target = a.placeholderTarget(["obj"]);
  const w = target ? target.w : p.width * 0.6;
  const h = target ? target.h : p.height * 0.6;
  const ch = makeChart(type, target ? target.x : (p.width - w) / 2, target ? target.y : (p.height - h) / 2, w, h, `Grafik ${a.nextShapeNumber()}`, grouping);
  const i = a.s.slideIdx;
  a.store.commit((pp) => updateSlideShapes(pp, i, (list) => (target ? list.map((s) => (s.id === target.id ? ch : s)) : [...list, ch])), { sel: [ch.id], editing: null });
  return ch.id;
}

export function updateChart(a: Actions, id: string, fn: (c: ChartData) => ChartData, merge?: string) {
  a.updateShapes((s) => (s.kind === "chart" ? { ...s, chart: fn(s.chart) } : s), [id], merge);
}

export function selectedChart(a: Actions): ChartShape | undefined {
  const s = a.selectedShapes();
  return s.length === 1 && s[0].kind === "chart" ? s[0] : undefined;
}

function readDataUrl(f: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result));
    r.onerror = () => reject(r.error);
    r.readAsDataURL(f);
  });
}

/** SVG/görüntü veri adresini PNG'ye çevirir (PPTX kapak resmi için). */
async function rasterize(src: string, w: number, h: number): Promise<string> {
  const img = new Image();
  await new Promise<void>((res, rej) => {
    img.onload = () => res();
    img.onerror = () => rej(new Error("kapak"));
    img.src = src;
  });
  const c = document.createElement("canvas");
  c.width = w;
  c.height = h;
  c.getContext("2d")!.drawImage(img, 0, 0, w, h);
  return c.toDataURL("image/png");
}

/** Videonun ilk karesi ve boyutu (yüklenemezse koyu kapak). */
async function videoPoster(src: string): Promise<{ poster: string; w: number; h: number }> {
  try {
    const v = document.createElement("video");
    v.muted = true;
    v.preload = "auto";
    v.src = src;
    await new Promise<void>((res, rej) => {
      v.onloadeddata = () => res();
      v.onerror = () => rej(new Error("video"));
      window.setTimeout(() => rej(new Error("zaman aşımı")), 4000);
    });
    const w = v.videoWidth || 640;
    const h = v.videoHeight || 360;
    try {
      v.currentTime = Math.min(0.5, (v.duration || 1) / 10);
      await new Promise<void>((res) => {
        v.onseeked = () => res();
        window.setTimeout(res, 1500);
      });
    } catch {
      // ilk kare kullanılır
    }
    const c = document.createElement("canvas");
    c.width = w;
    c.height = h;
    const ctx = c.getContext("2d")!;
    ctx.fillStyle = "#000";
    ctx.fillRect(0, 0, w, h);
    ctx.drawImage(v, 0, 0, w, h);
    return { poster: c.toDataURL("image/png"), w, h };
  } catch {
    const c = document.createElement("canvas");
    c.width = 640;
    c.height = 360;
    const ctx = c.getContext("2d")!;
    ctx.fillStyle = "#1f1f1f";
    ctx.fillRect(0, 0, 640, 360);
    ctx.fillStyle = "#fff";
    ctx.beginPath();
    ctx.moveTo(290, 140);
    ctx.lineTo(370, 180);
    ctx.lineTo(290, 220);
    ctx.closePath();
    ctx.fill();
    return { poster: c.toDataURL("image/png"), w: 640, h: 360 };
  }
}

/** Dosyadan video ya da ses ekler. */
export async function insertMediaFile(a: Actions, file: File, kind: "video" | "audio") {
  const mime = file.type || (kind === "video" ? "video/mp4" : "audio/mpeg");
  if (!mime.startsWith(kind === "video" ? "video/" : "audio/")) {
    a.message(kind === "video" ? "Seçilen dosya bir video değil." : "Seçilen dosya bir ses dosyası değil.", "warning");
    return;
  }
  if (file.size > 200 * 1024 * 1024) {
    a.message("Ortam dosyası çok büyük (en fazla 200 MB).", "warning");
    return;
  }
  const src = await readDataUrl(file);
  const media: MediaInfo = { kind, src, mime };
  const p = a.pres;
  let pic: PicShape;
  const name = file.name.replace(/\.[^.]+$/, "");
  if (kind === "video") {
    const { poster, w, h } = await videoPoster(src);
    // Boş içerik yer tutucusu varsa video oraya sığdırılır.
    const target = a.placeholderTarget(["obj"]);
    const box = target ?? { x: p.width * 0.2, y: p.height * 0.2, w: p.width * 0.6, h: p.height * 0.6 };
    const s = Math.min(target ? Infinity : 1, box.w / w, box.h / h);
    const ww = w * s;
    const hh = h * s;
    pic = { kind: "pic", id: uid(), name: name || `Video ${a.nextShapeNumber()}`, x: box.x + (box.w - ww) / 2, y: box.y + (box.h - hh) / 2, w: ww, h: hh, rot: 0, src: poster, line: { fill: { type: "none" }, width: 1, dash: "solid" }, media };
    if (target) {
      pic.ph = target.ph;
      const i = a.s.slideIdx;
      const np = pic;
      a.store.commit((pp) => updateSlideShapes(pp, i, (list) => list.map((x) => (x.id === target.id ? np : x))), { sel: [np.id], editing: null });
      return;
    }
  } else {
    const poster = await rasterize(SPEAKER_POSTER, 96, 96).catch(() => "");
    const sz = 48;
    pic = { kind: "pic", id: uid(), name: name || `Ses ${a.nextShapeNumber()}`, x: (p.width - sz) / 2, y: (p.height - sz) / 2, w: sz, h: sz, rot: 0, src: poster, line: { fill: { type: "none" }, width: 1, dash: "solid" }, media };
  }
  if (!a.slide()) a.newSlide();
  const i = a.s.slideIdx;
  a.store.commit((pp) => updateSlideShapes(pp, i, (list) => [...list, pic]), { sel: [pic.id], editing: null });
}

/** Ortam dosyası seçme penceresini açar. */
export function pickMedia(a: Actions, kind: "video" | "audio") {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = kind === "video" ? "video/*,.mp4,.webm,.mov,.m4v" : "audio/*,.mp3,.wav,.m4a,.ogg";
  input.onchange = () => {
    const f = input.files?.[0];
    if (f) void insertMediaFile(a, f, kind).catch(() => a.message("Ortam dosyası eklenemedi.", "error"));
  };
  input.click();
}

export function setMediaOption(a: Actions, patch: Partial<MediaInfo>) {
  a.updateShapes((s) => (s.kind === "pic" && s.media ? { ...s, media: { ...s.media, ...patch } } : s));
}
