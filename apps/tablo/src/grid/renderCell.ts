/** Hücre çizim yardımcıları: metin ölçümü, desen dolgusu, kenarlıklar, metin (kaydırma/taşma), filtre düğmesi. */
import { fontStack } from "@officehere/core";
import { formatGeneral } from "@officehere/formula";
import type { CellView } from "../model/display";
import { DEFAULT_SIZE, fontCss, type BorderSide, type CellStyle } from "../model/styles";

let measureCtx: CanvasRenderingContext2D | null = null;

export function measureText(text: string, font: string): number {
  if (!measureCtx && typeof document !== "undefined") measureCtx = document.createElement("canvas").getContext("2d");
  if (!measureCtx) {
    // DOM yok (birim testleri): yazı tipi boyutuna göre yaklaşık genişlik
    const px = Number(/([\d.]+)px/.exec(font)?.[1] ?? 14.67);
    return text.length * px * 0.5;
  }
  measureCtx.font = font;
  return measureCtx.measureText(text).width;
}

/** Hücre metni için kullanılabilir yazı tipi tanımı. */
export function cellFont(s: CellStyle, z: number): string {
  return fontCss(s, z, fontStack);
}

export function snap(v: number, dpr: number): number {
  return Math.round(v * dpr) / dpr;
}

export function drawPattern(ctx: CanvasRenderingContext2D, p: { type: string; color: string }, x: number, y: number, w: number, h: number): void {
  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();
  ctx.strokeStyle = p.color;
  ctx.lineWidth = 1;
  const step = p.type.includes("light") ? 6 : 4;
  ctx.beginPath();
  if (p.type.includes("Horizontal") || p.type.includes("Grid") || p.type.includes("Gray") || p.type === "gray125") for (let yy = y; yy < y + h; yy += step) {
    ctx.moveTo(x, yy + 0.5);
    ctx.lineTo(x + w, yy + 0.5);
  }
  if (p.type.includes("Vertical") || p.type.includes("Grid")) for (let xx = x; xx < x + w; xx += step) {
    ctx.moveTo(xx + 0.5, y);
    ctx.lineTo(xx + 0.5, y + h);
  }
  if (p.type.includes("Down") || p.type.includes("Trellis")) for (let k = -h; k < w; k += step) {
    ctx.moveTo(x + k, y);
    ctx.lineTo(x + k + h, y + h);
  }
  if (p.type.includes("Up") || p.type.includes("Trellis")) for (let k = 0; k < w + h; k += step) {
    ctx.moveTo(x + k, y);
    ctx.lineTo(x + k - h, y + h);
  }
  ctx.stroke();
  ctx.restore();
}

export function borderWidth(s: BorderSide["style"]): number {
  switch (s) {
    case "medium":
    case "mediumDashed":
    case "mediumDashDot":
    case "slantDashDot":
      return 2;
    case "thick":
      return 3;
    case "double":
      return 3;
    default:
      return 1;
  }
}

export function drawBorder(ctx: CanvasRenderingContext2D, b: BorderSide, x1: number, y1: number, x2: number, y2: number, dpr: number, dir: "h" | "v"): void {
  const color = b.color ?? "#000000";
  const w = borderWidth(b.style) / dpr;
  const px = 1 / dpr;
  ctx.fillStyle = color;
  const sx1 = snap(x1, dpr);
  const sy1 = snap(y1, dpr);
  const sx2 = snap(x2, dpr);
  const sy2 = snap(y2, dpr);
  if (b.style === "double") {
    if (dir === "h") {
      ctx.fillRect(sx1 - px, sy1 - 2 * px, sx2 - sx1 + px, px);
      ctx.fillRect(sx1 - px, sy1, sx2 - sx1 + px, px);
    } else {
      ctx.fillRect(sx1 - 2 * px, sy1 - px, px, sy2 - sy1 + px);
      ctx.fillRect(sx1, sy1 - px, px, sy2 - sy1 + px);
    }
    return;
  }
  const dashed = b.style.toLowerCase().includes("dash") || b.style === "dotted" || b.style === "hair";
  if (!dashed) {
    const off = Math.floor((borderWidth(b.style) - 1) / 2) / dpr + px;
    if (dir === "h") ctx.fillRect(sx1 - px, sy1 - off, sx2 - sx1 + px, w);
    else ctx.fillRect(sx1 - off, sy1 - px, w, sy2 - sy1 + px);
    return;
  }
  const pattern = b.style === "dotted" || b.style === "hair" ? [1, 1] : b.style.includes("DotDot") ? [3, 1, 1, 1, 1, 1] : b.style.includes("Dot") ? [3, 1, 1, 1] : [3, 1];
  const len = dir === "h" ? sx2 - sx1 : sy2 - sy1;
  let pos = 0;
  let k = 0;
  while (pos < len) {
    const seg = pattern[k % pattern.length] * (borderWidth(b.style) > 1 ? 2 : 1.5);
    if (k % 2 === 0) {
      const l = Math.min(seg, len - pos);
      if (dir === "h") ctx.fillRect(sx1 + pos, sy1 - px, l, w);
      else ctx.fillRect(sx1 - px, sy1 + pos, w, l);
    }
    pos += seg;
    k++;
  }
}

/** Kaydırmalı metin satırları. */
export function wrapLines(text: string, font: string, maxW: number): string[] {
  const out: string[] = [];
  for (const para of text.split("\n")) {
    const words = para.split(/(\s+)/);
    let line = "";
    for (const w of words) {
      const test = line + w;
      if (line && measureText(test.trimEnd(), font) > maxW) {
        out.push(line.trimEnd());
        line = w.trimStart();
        // Tek sözcük sığmıyorsa harf harf böl
        while (measureText(line, font) > maxW && line.length > 1) {
          let k = line.length - 1;
          while (k > 1 && measureText(line.slice(0, k), font) > maxW) k--;
          out.push(line.slice(0, k));
          line = line.slice(k);
        }
      } else line = test;
    }
    out.push(line.trimEnd());
  }
  return out;
}

export function drawCellText(ctx: CanvasRenderingContext2D, v: CellView, x: number, y: number, w: number, h: number, cellX: number, cellW: number, z: number, dpr: number): void {
  const s = v.style;
  let font = cellFont(s, z);
  const size = (s.size ?? DEFAULT_SIZE) * (96 / 72) * z;
  const pad = 3 * z;
  const indent = (s.indent ?? 0) * 9 * z;
  let text = v.text;
  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();
  ctx.font = font;
  ctx.fillStyle = v.color;
  const avail = cellW - pad * 2 - indent;
  let hashed = false;
  // Sayılar sığmazsa ### (Genel biçimde önce basamak azaltılır)
  if (v.isNumber && !s.wrap) {
    let tw = ctx.measureText(text).width;
    if (tw > avail && s.shrink) {
      const k = Math.max(0.3, avail / tw);
      font = fontCss({ ...s, size: (s.size ?? DEFAULT_SIZE) * k }, z, fontStack);
      ctx.font = font;
      tw = ctx.measureText(text).width;
    }
    if (tw > avail) {
      if ((!s.numFmt || s.numFmt === "General") && typeof v.value === "number") {
        let done = false;
        for (let p = 10; p >= 1 && !done; p--) {
          const t = shortGeneral(v.value, p);
          if (ctx.measureText(t).width <= avail) {
            text = t;
            done = true;
          }
        }
        if (!done) {
          text = "#".repeat(Math.max(1, Math.floor(avail / Math.max(1, ctx.measureText("#").width))));
          hashed = true;
        }
      } else {
        text = "#".repeat(Math.max(1, Math.floor(avail / Math.max(1, ctx.measureText("#").width))));
        hashed = true;
      }
    }
  }
  // Doldurma karakteri (muhasebe biçimindeki "* ") yalnız değer sığdığında uygulanır: sığmayan değer düz ######## olur
  if (v.fillChar && !s.wrap && !hashed) {
    const fw = ctx.measureText(v.fillChar.char).width || 3;
    const tw = ctx.measureText(text).width;
    const n = Math.max(0, Math.floor((avail - tw) / fw));
    text = text.slice(0, v.fillChar.index) + v.fillChar.char.repeat(n) + text.slice(v.fillChar.index);
  }
  const lineH = size * 1.2;
  const lines = s.wrap ? wrapLines(text, font, Math.max(4, cellW - pad * 2 - indent)) : text.split("\n").slice(0, 1).concat(text.includes("\n") && !s.wrap ? [] : []);
  const totalH = lines.length * lineH;
  let ty: number;
  const va = s.vAlign ?? "bottom";
  if (va === "top") ty = y + 2 * z + size * 0.95;
  else if (va === "middle" || va === "justify" || va === "distributed") ty = y + (h - totalH) / 2 + size * 0.95;
  else ty = y + h - totalH - 2 * z + size * 0.95 + (lineH - size) * 0.5;
  const rot = s.rotation ?? 0;
  if (rot && !s.wrap) {
    const angle = rot === 255 ? 0 : rot > 90 ? (rot - 90) * (Math.PI / 180) : -rot * (Math.PI / 180);
    ctx.translate(cellX + cellW / 2, y + h / 2);
    if (rot === 255) {
      // Dikey yığılmış metin
      ctx.textAlign = "center";
      const chars = text.split("");
      chars.forEach((ch, i) => ctx.fillText(ch, 0, -((chars.length - 1) * lineH) / 2 + i * lineH + size / 3));
    } else {
      ctx.rotate(angle);
      ctx.textAlign = "center";
      ctx.fillText(text, 0, size / 3);
    }
    ctx.restore();
    return;
  }
  for (const line of lines) {
    const tw = ctx.measureText(line).width;
    let tx: number;
    if (v.align === "left") tx = cellX + pad + indent;
    else if (v.align === "right") tx = cellX + cellW - pad - tw - indent;
    else tx = cellX + (cellW - tw) / 2;
    if (s.hAlign === "fill" && line) {
      const reps = Math.max(1, Math.floor((cellW - pad * 2) / Math.max(1, tw)));
      ctx.fillText(line.repeat(reps), cellX + pad, ty);
    } else ctx.fillText(line, tx, ty);
    if (s.underline || s.strike) {
      const lw = Math.max(1 / dpr, size / 14);
      if (s.underline) {
        ctx.fillRect(tx, ty + size * 0.12, tw, lw);
        if (s.underline === "double") ctx.fillRect(tx, ty + size * 0.12 + lw * 2, tw, lw);
      }
      if (s.strike) ctx.fillRect(tx, ty - size * 0.3, tw, lw);
    }
    ty += lineH;
  }
  ctx.restore();
}

export function shortGeneral(v: number, p: number): string {
  const abs = Math.abs(v);
  if (abs >= 1e11 || (abs < 1e-4 && abs > 0)) {
    const s = v.toExponential(Math.max(0, p - 1));
    const m = /^(-?)(\d)(?:\.(\d+))?e([+-])(\d+)$/.exec(s);
    if (!m) return s;
    const frac = (m[3] ?? "").replace(/0+$/, "");
    return `${m[1]}${m[2]}${frac ? "," + frac : ""}E${m[4]}${m[5].padStart(2, "0")}`;
  }
  const r = Number(v.toPrecision(p));
  return formatGeneral(r).replace(/^(-?\d+),(\d*?)0+$/, "$1,$2").replace(/,$/, "");
}

export function drawFilterButton(ctx: CanvasRenderingContext2D, x: number, y: number, s: number, active: boolean, dpr: number): void {
  const px = 1 / dpr;
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(x, y, s, s);
  ctx.strokeStyle = "#ABABAB";
  ctx.lineWidth = px;
  ctx.strokeRect(snap(x, dpr) + px / 2, snap(y, dpr) + px / 2, s - px, s - px);
  ctx.fillStyle = "#444444";
  const cx = x + s / 2;
  const cy = y + s / 2;
  const a = s * 0.18;
  if (active) {
    // Huni simgesi
    ctx.beginPath();
    ctx.moveTo(cx - a * 1.6, cy - a * 1.2);
    ctx.lineTo(cx + a * 1.6, cy - a * 1.2);
    ctx.lineTo(cx + a * 0.3, cy + a * 0.2);
    ctx.lineTo(cx + a * 0.3, cy + a * 1.5);
    ctx.lineTo(cx - a * 0.3, cy + a * 1.1);
    ctx.lineTo(cx - a * 0.3, cy + a * 0.2);
    ctx.closePath();
    ctx.fill();
  } else {
    ctx.beginPath();
    ctx.moveTo(cx - a * 1.3, cy - a * 0.5);
    ctx.lineTo(cx + a * 1.3, cy - a * 0.5);
    ctx.lineTo(cx, cy + a * 0.9);
    ctx.closePath();
    ctx.fill();
  }
}

/* ------------------------------------------------------------------ */
/* Seçim ve vurgu katmanları                                            */
/* ------------------------------------------------------------------ */

/** Hücre içi onay kutusunun dikdörtgeni: yatay hizalamaya göre (varsayılan ortada), dikeyde ortada. */
export function checkboxRect(cell: { x: number; y: number; w: number; h: number }, s: CellStyle, z: number): { x: number; y: number; s: number } {
  const size = Math.max(6, Math.min(13 * z, cell.h - 4 * z, cell.w - 4 * z));
  const pad = 3 * z + (s.indent ?? 0) * 9 * z;
  const x = s.hAlign === "left" ? cell.x + pad : s.hAlign === "right" ? cell.x + cell.w - pad - size : cell.x + (cell.w - size) / 2;
  const y = s.vAlign === "top" ? cell.y + 2 * z : s.vAlign === "bottom" ? cell.y + cell.h - size - 2 * z : cell.y + (cell.h - size) / 2;
  return { x, y, s: size };
}

/** Onay kutusunu çizer: boşken yazı rengiyle çerçeve, işaretliyken yazı rengiyle dolu kutuda beyaz onay işareti. */
export function drawCheckbox(ctx: CanvasRenderingContext2D, x: number, y: number, s: number, checked: boolean, color: string, dpr: number): void {
  const px = 1 / dpr;
  const r = Math.max(1.5, s * 0.18);
  const bx = snap(x, dpr) + px / 2;
  const by = snap(y, dpr) + px / 2;
  const bs = Math.round(s * dpr) / dpr - px;
  ctx.save();
  ctx.beginPath();
  ctx.roundRect(bx, by, bs, bs, r);
  if (checked) {
    ctx.fillStyle = color;
    ctx.fill();
    ctx.strokeStyle = "#FFFFFF";
    ctx.lineWidth = Math.max(1.5, s * 0.14);
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    ctx.beginPath();
    ctx.moveTo(bx + bs * 0.24, by + bs * 0.52);
    ctx.lineTo(bx + bs * 0.43, by + bs * 0.72);
    ctx.lineTo(bx + bs * 0.77, by + bs * 0.3);
    ctx.stroke();
  } else {
    ctx.fillStyle = "#FFFFFF";
    ctx.fill();
    ctx.strokeStyle = color;
    ctx.lineWidth = Math.max(px, s / 13);
    ctx.stroke();
  }
  ctx.restore();
}
