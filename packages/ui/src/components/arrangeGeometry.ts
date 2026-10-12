/**
 * Yerleşim çizimlerinin geometrisi: Metni Kaydır, Konum (sayfa önizlemesi), Öne Getir / Arkaya Gönder
 * ve Hizala. Her çizim istenen piksel boyutunda tamsayı koordinatlarla yeniden kurulur; düz kenarlar
 * 16, 20, 24, 32 px'te ve galeri önizlemelerinde piksel ızgarasına oturur (ölçeklenip bulanıklaşmaz).
 * DOM'a bağlı değil; çizim ArrangeIcons.tsx'te.
 */

/** Boya rolü: çizim teması bu rollere renk verir (gri satır, vurgu rengindeki resim, sayfa zemini...). */
export type Ink =
  | "line" // metin satırı (nötr gri)
  | "pic" // resim / nesne (vurgu rengi)
  | "ghost" // metnin arkasındaki resim (soluk vurgu)
  | "mark" // resmin üstündeki dağ ve güneş işareti
  | "page" // sayfa ya da kart zemini
  | "edge" // sayfa çerçevesi (ince, yarı saydam)
  | "ink"; // kılavuz çizgisi, kart kenarı (simge rengi)

export type Prim =
  | { k: "rect"; x: number; y: number; w: number; h: number; ink: Ink }
  /** İçi boş çerçeve (dolgu kuralı evenodd; kenar kalınlığı sw). */
  | { k: "frame"; x: number; y: number; w: number; h: number; sw: number; ink: Ink }
  | { k: "path"; d: string; ink: Ink };

export interface Drawing {
  w: number;
  h: number;
  prims: Prim[];
}

export type WrapKind = "inline" | "square" | "tight" | "topBottom" | "behind" | "front";
export type LayerKind = "forward" | "front" | "backward" | "back";
export type AlignKind = "left" | "center" | "right" | "top" | "middle" | "bottom" | "distH" | "distV";
export type PositionKey = "inline" | { v: "top" | "middle" | "bottom"; h: "left" | "center" | "right" };

const R = Math.round;

function rect(x: number, y: number, w: number, h: number, ink: Ink): Prim {
  return { k: "rect", x, y, w, h, ink };
}

/** Metin satırı parçası; boyu bir pikselden kısaysa çizilmez. */
function seg(out: Prim[], x0: number, x1: number, y: number, lw: number) {
  if (x1 - x0 >= 1) out.push(rect(x0, y, x1 - x0, lw, "line"));
}

/** Resim: vurgu renginde dikdörtgen; yeterince büyükse üstünde küçük dağ ve güneş işareti. */
function picture(out: Prim[], x: number, y: number, w: number, h: number, ink: "pic" | "ghost" = "pic") {
  out.push(rect(x, y, w, h, ink));
  if (ink !== "pic" || w < 10 || h < 8) return;
  const f = (n: number) => +n.toFixed(2);
  const b = y + h;
  const mountain = `M${f(x + w * 0.08)} ${b}L${f(x + w * 0.38)} ${f(y + h * 0.42)}L${f(x + w * 0.6)} ${f(y + h * 0.72)}L${f(x + w * 0.72)} ${f(y + h * 0.58)}L${f(x + w * 0.92)} ${b}Z`;
  const r = Math.max(1, h * 0.12);
  const cx = x + w * 0.74;
  const cy = y + h * 0.28;
  const sun = `M${f(cx - r)} ${f(cy)}a${f(r)} ${f(r)} 0 1 0 ${f(2 * r)} 0a${f(r)} ${f(r)} 0 1 0 ${f(-2 * r)} 0Z`;
  out.push({ k: "path", d: mountain, ink: "mark" }, { k: "path", d: sun, ink: "mark" });
}

/* ------------------------------------------------------------------ */
/* Metni Kaydır                                                        */
/* ------------------------------------------------------------------ */

/** Kare simgede beş satırlık metin ızgarası: kenar boşluğu, satır kalınlığı, satır aralığı. */
export function textRows(s: number) {
  const pad = Math.max(1, R(s / 16));
  const lw = Math.max(1, Math.floor(s / 16));
  const pitch = Math.floor((s - 2 * pad - lw) / 4);
  const top = pad + Math.floor((s - 2 * pad - (4 * pitch + lw)) / 2);
  return { pad, lw, pitch, ys: [0, 1, 2, 3, 4].map((i) => top + i * pitch), x0: pad, x1: s - pad };
}

/**
 * Metni Kaydır seçeneğinin öğretici çizimi: resim 2.–4. satırların yüksekliğinde.
 * Kare: resim solda, satırlar sağında eşit boşlukla kısalır. Sıkı: eşkenar dörtgen biçimli resmin
 * çevresini satırlar yakından izler. Üst ve Alt: resmin yanı boş. Metinle Aynı Hizada: resim bir satırın
 * içinde (yanında aynı satırın parçaları). Metnin Arkasında: satırlar soluk resmin üstünden geçer.
 * Metnin Önünde: opak resim satırların üstünde.
 */
export function wrapDrawing(mode: WrapKind, s: number): Drawing {
  const { lw, pitch, ys, x0, x1 } = textRows(s);
  const out: Prim[] = [];
  const span = x1 - x0;
  let pw = R(span * 0.45);
  if ((span - pw) % 2) pw += 1;
  const ph = 2 * pitch + lw;
  const py = ys[1];
  const cx = x0 + (span - pw) / 2;
  const full = (i: number) => seg(out, x0, x1, ys[i], lw);
  switch (mode) {
    case "inline": {
      const gap = Math.max(1, Math.floor(s / 16));
      [0, 4].forEach(full);
      seg(out, x0, cx - gap, ys[3], lw);
      seg(out, cx + pw + gap, x1, ys[3], lw);
      picture(out, cx, py, pw, ph);
      break;
    }
    case "square": {
      const gap = Math.max(2, R(s / 8));
      [0, 4].forEach(full);
      for (const i of [1, 2, 3]) seg(out, x0 + pw + gap, x1, ys[i], lw);
      picture(out, x0, py, pw, ph);
      break;
    }
    case "tight": {
      // Eşkenar dörtgen: satır bandının merkeze en yakın noktasındaki yarı genişliğe göre satır başlar.
      const gap = Math.max(1, Math.floor(s / 16));
      const rx = pw / 2;
      const ry = ph / 2;
      const mx = x0 + rx;
      const my = py + ry;
      [0, 4].forEach(full);
      for (const i of [1, 2, 3]) {
        const a = ys[i];
        const b = a + lw;
        const dy = a <= my && my <= b ? 0 : Math.min(Math.abs(a - my), Math.abs(b - my));
        const half = rx * Math.max(0, 1 - dy / ry);
        seg(out, Math.ceil(mx + half + gap), x1, ys[i], lw);
      }
      out.push({ k: "path", d: `M${mx} ${py}L${x0 + pw} ${my}L${mx} ${py + ph}L${x0} ${my}Z`, ink: "pic" });
      break;
    }
    case "topBottom":
      [0, 4].forEach(full);
      picture(out, cx, py, pw, ph);
      break;
    case "behind":
      picture(out, cx, py, pw, ph, "ghost");
      [0, 1, 2, 3, 4].forEach(full);
      break;
    case "front":
      [0, 1, 2, 3, 4].forEach(full);
      picture(out, cx, py, pw, ph);
      break;
  }
  return { w: s, h: s, prims: out };
}

/* ------------------------------------------------------------------ */
/* Konum: sayfa önizlemesi                                             */
/* ------------------------------------------------------------------ */

/**
 * Sayfa önizlemesi (W×H px): ince çerçeveli sayfa, gri metin satırları ve konumdaki resim. Sol/sağ
 * konumlarda satırlar resmin yanında kısalır (kare kaydırma); ortadaki sütunda resmin yanı boştur
 * (üst ve alt kaydırma: düzenleyici ortalanan resmi böyle yerleştirir). "inline": resim bir satırın içinde.
 */
export function positionDrawing(pos: PositionKey, W: number, H: number): Drawing {
  const out: Prim[] = [rect(0, 0, W, H, "page"), { k: "frame", x: 0, y: 0, w: W, h: H, sw: 1, ink: "edge" }];
  const x0 = R(W * 0.125);
  const x1 = W - x0;
  const y0 = R(H * 0.1);
  const y1 = H - y0;
  const lw = Math.max(1, R(W / 20));
  const pitch = lw + Math.max(2, R(W / 20));
  const ys: number[] = [];
  for (let y = y0; y + lw <= y1; y += pitch) ys.push(y);
  const n = ys.length;
  const span = x1 - x0;
  const pw = 2 * R(span * 0.21);
  const ph = 2 * pitch + lw;
  const gap = Math.max(2, R(W / 13));
  // Paragraf sonu satırları kısadır (metin görünümü).
  const shortEnd = x0 + R(span * 0.62);
  const lineEnd = (i: number) => (i % 4 === 3 || i === n - 1 ? shortEnd : x1);

  let k: number;
  let px: number;
  if (pos === "inline") {
    k = Math.min(3, n - 3);
    px = x0 + (span - pw) / 2;
  } else {
    k = pos.v === "top" ? 0 : pos.v === "bottom" ? n - 3 : Math.floor((n - 3) / 2);
    px = pos.h === "left" ? x0 : pos.h === "right" ? x1 - pw : x0 + (span - pw) / 2;
  }
  for (let i = 0; i < n; i++) {
    const beside = i >= k && i <= k + 2;
    if (!beside) {
      seg(out, x0, lineEnd(i), ys[i], lw);
      continue;
    }
    if (pos === "inline") {
      // Resim satırın içinde: yalnız taban satırında (en alttaki) yanında metin parçaları.
      if (i === k + 2) {
        seg(out, x0, px - lw, ys[i], lw);
        seg(out, px + pw + lw, x1, ys[i], lw);
      }
    } else if (pos.h === "left") seg(out, px + pw + gap, x1, ys[i], lw);
    else if (pos.h === "right") seg(out, x0, px - gap, ys[i], lw);
  }
  picture(out, px, ys[k], pw, ph);
  return { w: W, h: H, prims: out };
}

/* ------------------------------------------------------------------ */
/* Öne Getir / Arkaya Gönder                                            */
/* ------------------------------------------------------------------ */

/**
 * Katman sırası: çapraz dizilmiş kartlar, sağ alttaki en öndedir. Bir düzey (Öne Getir / Arkaya Gönder)
 * iki kartla, en öne / en arkaya üç kartla gösterilir; vurgu renkli kart taşınan nesnedir.
 */
export function layerDrawing(kind: LayerKind, s: number): Drawing {
  const pad = Math.max(1, R(s / 16));
  const sw = Math.max(1, Math.floor(s / 16));
  const three = kind === "front" || kind === "back";
  const c = three ? R(s * 0.5) : R(s * 0.62);
  const room = s - 2 * pad - c;
  const steps = three ? [0, Math.floor(room / 2), room] : [0, room];
  const target = kind === "front" || kind === "forward" ? steps.length - 1 : 0;
  const out: Prim[] = [];
  steps.forEach((o, i) => {
    const x = pad + o;
    if (i === target) out.push(rect(x, x, c, c, "pic"));
    else out.push(rect(x, x, c, c, "page"), { k: "frame", x, y: x, w: c, h: c, sw, ink: "ink" });
  });
  return { w: s, h: s, prims: out };
}

/* ------------------------------------------------------------------ */
/* Hizala                                                              */
/* ------------------------------------------------------------------ */

/** Yatay çizimi köşegene göre yansıtır (sola/ortaya/sağa → üste/ortaya/alta). */
function transpose(p: Prim): Prim {
  if (p.k === "rect") return { ...p, x: p.y, y: p.x, w: p.h, h: p.w };
  if (p.k === "frame") return { ...p, x: p.y, y: p.x, w: p.h, h: p.w };
  return p;
}

/**
 * Hizalama: iki nesne (uzun ve kısa çubuk, vurgu renginde) ve hizalandıkları kılavuz çizgisi.
 * Dağıt: eşit aralıklı, farklı boyda üç nesne.
 */
export function alignDrawing(kind: AlignKind, s: number): Drawing {
  const pad = Math.max(1, R(s / 16));
  const sw = Math.max(1, Math.floor(s / 16));
  const vertical = kind === "top" || kind === "middle" || kind === "bottom" || kind === "distV";
  const out: Prim[] = [];
  if (kind === "distH" || kind === "distV") {
    const avail = s - 2 * pad;
    const bw = R(avail * 0.2);
    const gap = Math.floor((avail - 3 * bw) / 2);
    const x = pad + Math.floor((avail - 3 * bw - 2 * gap) / 2);
    const heights = [0.6, 1, 0.45].map((f) => {
      const h = Math.max(2, R(avail * f));
      return h + ((avail - h) % 2);
    });
    heights.forEach((h, i) => out.push(rect(x + i * (bw + gap), pad + (avail - h) / 2, bw, h, "pic")));
  } else {
    const t = R(s * 0.25);
    const g = Math.max(1, R(s / 8));
    const y = Math.floor((s - (2 * t + g)) / 2);
    let long = R(s * 0.62);
    let short = R(s * 0.38);
    const side = kind === "left" || kind === "top" ? "start" : kind === "right" || kind === "bottom" ? "end" : "center";
    let gx: number;
    let starts: [number, number];
    if (side === "center") {
      // Kılavuz ortada; çubuklar ona göre tam ortalansın diye uzunluklar kılavuz kalınlığıyla aynı teklikte.
      gx = Math.floor((s - sw) / 2);
      long += (long - sw) % 2;
      short += (short - sw) % 2;
      starts = [gx + (sw - long) / 2, gx + (sw - short) / 2];
    } else if (side === "start") {
      gx = pad;
      starts = [gx + sw + sw, gx + sw + sw];
    } else {
      gx = s - pad - sw;
      starts = [gx - sw - long, gx - sw - short];
    }
    out.push(rect(starts[0], y, long, t, "pic"), rect(starts[1], y + t + g, short, t, "pic"));
    out.push(rect(gx, pad, sw, s - 2 * pad, "ink"));
  }
  return { w: s, h: s, prims: vertical ? out.map(transpose) : out };
}
