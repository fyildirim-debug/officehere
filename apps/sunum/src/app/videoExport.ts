/**
 * Videoya aktarma: gizli olmayan slaytlar sırayla, geçişleri ve animasyon
 * adımlarıyla (adımlar arası yumuşak geçiş) bir WebM videosuna çizilir.
 * Kareler tuvalde oluşturulur, WebCodecs VideoEncoder (VP9/VP8) ile
 * kodlanır ve io/webm ile dosyaya yazılır; gerçek zamanlı kayıt gerekmez.
 * Durağan bölümler saniyede bir anahtar kareyle (değişken kare hızı) yazılır.
 */
import type { Presentation, Slide, Transition } from "../model/types";
import { buildSteps, initiallyHidden } from "../show/effects";
import { WebmMuxer } from "../io/webm";
import { slideToCanvas } from "./exporting";

export interface VideoOptions {
  /** Video yüksekliği (piksel); genişlik slayt oranından. */
  height: 1080 | 720 | 480;
  /** Zamanlaması olmayan slaytlarda geçirilecek saniye. */
  secondsPerSlide: number;
  /** Kaydedilmiş slayt zamanlamalarını (otomatik ilerleme) kullan. */
  useTimings: boolean;
}

export const DEFAULT_VIDEO_OPTIONS: VideoOptions = { height: 1080, secondsPerSlide: 5, useTimings: true };

/** Zaman çizelgesi parçası (ms). */
export type VideoSeg =
  | { kind: "hold"; slide: number; state: number; start: number; dur: number }
  | { kind: "trans"; slide: number; prev: { slide: number; state: number } | null; transition: Transition; start: number; dur: number }
  | { kind: "step"; slide: number; state: number; start: number; dur: number };

const FPS = 30;
/** Durağan bölümlerde anahtar kare aralığı (ms). */
const KEY_EVERY = 1000;

/**
 * Slaydın animasyon durumları: 0 = otomatik efektlerden sonra, k = k. tıklamadan sonra.
 * Her durum o anda gizli şekillerin kümesidir (giriş efekti oynamamış ya da çıkmış).
 */
export function slideStates(slide: Slide): Set<string>[] {
  const st = buildSteps(slide);
  const hidden = initiallyHidden(slide);
  const apply = (effects: typeof st.auto) => {
    for (const e of [...effects].sort((a, b) => a.start - b.start)) {
      if (e.anim.cls === "entr") hidden.delete(e.anim.shapeId);
      else if (e.anim.cls === "exit") hidden.add(e.anim.shapeId);
    }
  };
  apply(st.auto);
  const out = [new Set(hidden)];
  for (const g of st.clicks) {
    apply(g);
    out.push(new Set(hidden));
  }
  return out;
}

/** Slaytta geçirilecek süre (ms). */
export function slideDuration(slide: Slide, opts: VideoOptions): number {
  const adv = slide.transition?.advanceAfter;
  return Math.max(0, (opts.useTimings && adv !== undefined ? adv : opts.secondsPerSlide) * 1000);
}

const hasTransition = (t: Transition | undefined): t is Transition => !!t && t.type !== "none" && t.type !== "cut";

/** Videonun zaman çizelgesi ve toplam süresi (ms). */
export function videoTimeline(pres: Presentation, opts: VideoOptions): { segs: VideoSeg[]; total: number } {
  const segs: VideoSeg[] = [];
  let t = 0;
  let prev: { slide: number; state: number } | null = null;
  pres.slides.forEach((sl, i) => {
    if (sl.hidden) return;
    const n = slideStates(sl).length;
    if (hasTransition(sl.transition)) {
      const dur = Math.max(100, Math.round(sl.transition.duration * 1000));
      segs.push({ kind: "trans", slide: i, prev, transition: sl.transition, start: t, dur });
      t += dur;
    }
    const per = Math.max(200, Math.round(slideDuration(sl, opts) / n));
    segs.push({ kind: "hold", slide: i, state: 0, start: t, dur: per });
    t += per;
    for (let s = 1; s < n; s++) {
      const cf = Math.min(500, Math.round(per / 2));
      segs.push({ kind: "step", slide: i, state: s, start: t, dur: cf });
      t += cf;
      segs.push({ kind: "hold", slide: i, state: s, start: t, dur: per - cf });
      t += per - cf;
    }
    prev = { slide: i, state: n - 1 };
  });
  return { segs, total: t };
}

/** Video boyutu: yükseklik verilir, genişlik slayt oranından (çift sayı). */
export function videoSize(pres: Presentation, height: number): { w: number; h: number } {
  return { w: Math.max(2, Math.round((height * pres.width) / pres.height / 2) * 2), h: height };
}

/* ------------------------- Geçiş çizimi ------------------------- */

type Img = CanvasImageSource;
const easeInOut = (p: number) => (p < 0.5 ? 2 * p * p : 1 - (-2 * p + 2) ** 2 / 2);
const DIR_VEC: Record<string, [number, number]> = { u: [0, 1], d: [0, -1], l: [1, 0], r: [-1, 0] };
/** Tekrarlanabilir sözde rastgele (0..1). */
const hash = (n: number) => {
  const x = Math.sin(n * 12.9898) * 43758.5453;
  return x - Math.floor(x);
};

/** Geçişin p ilerlemesindeki kare (from null: siyahtan). */
export function drawTransition(ctx: CanvasRenderingContext2D, from: Img | null, to: Img, p: number, tr: Transition, W: number, H: number) {
  const o = tr.option ?? "";
  const e = easeInOut(p);
  const draw = (img: Img | null, x = 0, y = 0, w = W, h = H) => {
    if (img) ctx.drawImage(img, x, y, w, h);
    else {
      ctx.fillStyle = "#000";
      ctx.fillRect(x, y, w, h);
    }
  };
  const clipped = (path: () => void) => {
    ctx.save();
    ctx.beginPath();
    path();
    ctx.clip();
    draw(to);
    ctx.restore();
  };
  ctx.globalAlpha = 1;
  switch (tr.type) {
    case "fade":
    case "reveal": {
      if (o === "black" || tr.type === "reveal") {
        ctx.fillStyle = "#000";
        ctx.fillRect(0, 0, W, H);
        ctx.globalAlpha = p < 0.5 ? 1 - p * 2 : p * 2 - 1;
        draw(p < 0.5 ? from : to);
      } else {
        draw(from);
        ctx.globalAlpha = p;
        draw(to);
      }
      break;
    }
    case "push": {
      const [vx, vy] = DIR_VEC[o || "u"];
      draw(from, -vx * e * W, -vy * e * H);
      draw(to, vx * (1 - e) * W, vy * (1 - e) * H);
      break;
    }
    case "cover": {
      const [vx, vy] = DIR_VEC[o || "l"];
      draw(from);
      draw(to, vx * (1 - e) * W, vy * (1 - e) * H);
      break;
    }
    case "uncover": {
      const [vx, vy] = DIR_VEC[o || "l"];
      draw(to);
      draw(from, -vx * e * W, -vy * e * H);
      break;
    }
    case "wipe": {
      draw(from);
      // Seçenek, efektin geldiği yön: l sağdan, r soldan, d üstten, u alttan.
      const d = o || "l";
      clipped(() => (d === "r" ? ctx.rect(0, 0, p * W, H) : d === "l" ? ctx.rect(W - p * W, 0, p * W, H) : d === "d" ? ctx.rect(0, 0, W, p * H) : ctx.rect(0, H - p * H, W, p * H)));
      break;
    }
    case "split": {
      draw(from);
      const v = o || "vertOut";
      const horz = v.startsWith("horz");
      const out = v.endsWith("Out");
      clipped(() => {
        if (horz && out) ctx.rect(0, H / 2 - (e * H) / 2, W, e * H);
        else if (horz) (ctx.rect(0, 0, W, (e * H) / 2), ctx.rect(0, H - (e * H) / 2, W, (e * H) / 2));
        else if (out) ctx.rect(W / 2 - (e * W) / 2, 0, e * W, H);
        else (ctx.rect(0, 0, (e * W) / 2, H), ctx.rect(W - (e * W) / 2, 0, (e * W) / 2, H));
      });
      break;
    }
    case "shape": {
      draw(from);
      const r = (e * Math.hypot(W, H)) / 2;
      clipped(() => {
        if (o === "diamond") {
          ctx.moveTo(W / 2, H / 2 - r * 1.42);
          ctx.lineTo(W / 2 + r * 1.42, H / 2);
          ctx.lineTo(W / 2, H / 2 + r * 1.42);
          ctx.lineTo(W / 2 - r * 1.42, H / 2);
          ctx.closePath();
        } else if (o === "plus") {
          ctx.rect(W / 2 - r * 0.7, 0, r * 1.4, H);
          ctx.rect(0, H / 2 - r * 0.7, W, r * 1.4);
        } else ctx.arc(W / 2, H / 2, r, 0, Math.PI * 2);
      });
      break;
    }
    case "zoom": {
      if (o === "out") {
        draw(to);
        const s = 1 + 0.8 * e;
        ctx.globalAlpha = 1 - p;
        draw(from, (W - W * s) / 2, (H - H * s) / 2, W * s, H * s);
      } else {
        draw(from);
        const s = 0.3 + 0.7 * e;
        ctx.globalAlpha = p;
        draw(to, (W - W * s) / 2, (H - H * s) / 2, W * s, H * s);
      }
      break;
    }
    case "wheel":
    case "wedge": {
      draw(from);
      clipped(() => {
        ctx.moveTo(W / 2, H / 2);
        const a0 = -Math.PI / 2;
        if (tr.type === "wheel") ctx.arc(W / 2, H / 2, Math.hypot(W, H), a0, a0 + p * Math.PI * 2);
        else ctx.arc(W / 2, H / 2, Math.hypot(W, H), a0 - p * Math.PI, a0 + p * Math.PI);
        ctx.closePath();
      });
      break;
    }
    case "blinds": {
      draw(from);
      const n = 8;
      clipped(() => {
        for (let i = 0; i < n; i++) {
          if (o === "horz") ctx.rect(0, (i * H) / n, W, (p * H) / n);
          else ctx.rect((i * W) / n, 0, (p * W) / n, H);
        }
      });
      break;
    }
    case "checker": {
      draw(from);
      const cols = 12;
      const sz = W / cols;
      const rows = Math.ceil(H / sz);
      clipped(() => {
        for (let r = 0; r < rows; r++)
          for (let c = 0; c < cols; c++) {
            const off = (r + c) % 2 ? 0.5 : 0;
            const q = Math.max(0, Math.min(1, p * 2 - off));
            if (q > 0) (o === "vert" ? ctx.rect(c * sz, r * sz, sz, sz * q) : ctx.rect(c * sz, r * sz, sz * q, sz));
          }
      });
      break;
    }
    case "dissolve":
    case "randomBar": {
      draw(from);
      clipped(() => {
        if (tr.type === "randomBar") {
          const n = 64;
          for (let i = 0; i < n; i++) if (hash(i + 1) < p) (o === "horz" ? ctx.rect(0, (i * H) / n, W, H / n + 0.5) : ctx.rect((i * W) / n, 0, W / n + 0.5, H));
        } else {
          const cell = Math.max(8, Math.round(W / 64));
          for (let y = 0, k = 0; y < H; y += cell) for (let x = 0; x < W; x += cell, k++) if (hash(k + 7) < p) ctx.rect(x, y, cell, cell);
        }
      });
      break;
    }
    case "newsflash": {
      draw(from);
      ctx.save();
      ctx.translate(W / 2, H / 2);
      ctx.rotate((1 - e) * Math.PI * 4);
      ctx.scale(Math.max(0.001, e), Math.max(0.001, e));
      draw(to, -W / 2, -H / 2);
      ctx.restore();
      break;
    }
    default:
      // Dönüştür ve diğerleri: yumuşak geçiş.
      draw(from);
      ctx.globalAlpha = p;
      draw(to);
  }
  ctx.globalAlpha = 1;
}

/* ------------------------- Kodlama ------------------------- */

export class VideoUnsupportedError extends Error {}

const BITRATE: Record<number, number> = { 1080: 8_000_000, 720: 5_000_000, 480: 2_500_000 };

async function pickCodec(w: number, h: number, bitrate: number): Promise<{ codec: string; id: string } | null> {
  if (typeof VideoEncoder === "undefined" || typeof VideoFrame === "undefined") return null;
  for (const c of [
    { codec: "vp09.00.10.08", id: "V_VP9" },
    { codec: "vp8", id: "V_VP8" },
  ]) {
    const r = await VideoEncoder.isConfigSupported({ codec: c.codec, width: w, height: h, bitrate, framerate: FPS, latencyMode: "realtime" }).catch(() => null);
    if (r?.supported) return c;
  }
  return null;
}

/**
 * Sunuyu WebM videosuna dönüştürür. onProgress 0..1; signal iptal eder
 * (AbortError fırlatılır). Kodlayıcı yoksa VideoUnsupportedError.
 */
export async function exportVideo(pres: Presentation, opts: VideoOptions, onProgress: (p: number) => void, signal?: AbortSignal): Promise<Blob> {
  const { w: W, h: H } = videoSize(pres, opts.height);
  const bitrate = BITRATE[opts.height] ?? 5_000_000;
  const codec = await pickCodec(W, H, bitrate);
  if (!codec) throw new VideoUnsupportedError("Bu ortamda video kodlayıcı (VP9/VP8) yok.");
  const { segs, total } = videoTimeline(pres, opts);
  if (!segs.length) throw new Error("Videoya aktarılacak slayt yok (tüm slaytlar gizli).");

  const mux = new WebmMuxer(W, H, codec.id);
  let failure: Error | null = null;
  const enc = new VideoEncoder({
    output: (chunk) => {
      const buf = new Uint8Array(chunk.byteLength);
      chunk.copyTo(buf);
      mux.add(buf, chunk.timestamp / 1000, chunk.type === "key");
    },
    error: (e) => (failure = e instanceof Error ? e : new Error(String(e))),
  });
  enc.configure({ codec: codec.codec, width: W, height: H, bitrate, framerate: FPS, latencyMode: "realtime" });

  const canvas = document.createElement("canvas");
  canvas.width = W;
  canvas.height = H;
  const ctx = canvas.getContext("2d")!;

  // Durum resimleri: yalnız geçerli ve bir önceki slaydınkiler bellekte tutulur.
  const cache = new Map<string, HTMLCanvasElement>();
  const statesOf = new Map<number, Set<string>[]>();
  const image = async (slide: number, state: number): Promise<HTMLCanvasElement> => {
    const key = `${slide}:${state}`;
    const hit = cache.get(key);
    if (hit) return hit;
    let states = statesOf.get(slide);
    if (!states) statesOf.set(slide, (states = slideStates(pres.slides[slide])));
    const hidden = states[Math.min(state, states.length - 1)];
    const c = await slideToCanvas(pres, slide, W / pres.width, (s) => (hidden.has(s.id) ? "sh-vhidden" : undefined));
    cache.set(key, c);
    return c;
  };
  const forget = (keep: number[]) => {
    for (const k of [...cache.keys()]) if (!keep.includes(Number(k.split(":")[0]))) cache.delete(k);
  };

  let lastKey = -Infinity;
  let lastTime = -1;
  const emit = async (time: number, key: boolean) => {
    if (signal?.aborted) throw new DOMException("İptal edildi", "AbortError");
    if (failure) throw failure;
    const t = Math.max(lastTime + 1, Math.round(time));
    lastTime = t;
    const k = key || t - lastKey >= KEY_EVERY;
    if (k) lastKey = t;
    const frame = new VideoFrame(canvas, { timestamp: t * 1000 });
    enc.encode(frame, { keyFrame: k });
    frame.close();
    while (enc.encodeQueueSize > 6) await new Promise((r) => setTimeout(r, 5));
    onProgress(Math.min(0.99, t / Math.max(1, total)));
  };

  try {
    let lastImg: HTMLCanvasElement | null = null;
    for (const seg of segs) {
      forget([seg.slide, seg.kind === "trans" && seg.prev ? seg.prev.slide : seg.slide]);
      const to = await image(seg.slide, seg.kind === "trans" ? 0 : seg.state);
      if (seg.kind === "hold") {
        ctx.drawImage(to, 0, 0, W, H);
        for (let t = seg.start; t < seg.start + seg.dur; t += KEY_EVERY) await emit(t, true);
      } else {
        const from = seg.kind === "trans" ? (seg.prev ? await image(seg.prev.slide, seg.prev.state) : null) : await image(seg.slide, seg.state - 1);
        const n = Math.max(1, Math.round((seg.dur * FPS) / 1000));
        for (let f = 0; f < n; f++) {
          const p = (f + 1) / n;
          ctx.clearRect(0, 0, W, H);
          if (seg.kind === "trans") drawTransition(ctx, from, to, p, seg.transition, W, H);
          else {
            ctx.drawImage(from!, 0, 0, W, H);
            ctx.globalAlpha = p;
            ctx.drawImage(to, 0, 0, W, H);
            ctx.globalAlpha = 1;
          }
          await emit(seg.start + (f * seg.dur) / n, f === 0);
        }
      }
      lastImg = to;
    }
    // Son kare videonun sonunda da görünsün.
    if (lastImg) ctx.drawImage(lastImg, 0, 0, W, H);
    await emit(total, true);
    await enc.flush();
  } finally {
    if (enc.state !== "closed") enc.close();
    cache.clear();
  }
  if (failure) throw failure;
  onProgress(1);
  return new Blob([mux.finish(total) as BlobPart], { type: "video/webm" });
}
