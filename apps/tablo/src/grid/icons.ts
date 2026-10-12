/** Koşullu biçimlendirme simge kümeleri: basit vektör çizimleri (oklar, ışıklar, işaretler, derecelendirmeler). */

const RED = "#E0412F";
const YELLOW = "#F2B705";
const GREEN = "#3A9E4A";
const GRAY = "#7F7F7F";
const BLACK = "#404040";

type Ctx = CanvasRenderingContext2D;

function arrow(ctx: Ctx, x: number, y: number, s: number, angleDeg: number, color: string): void {
  // Sağa bakan ok çizilip döndürülür (0: sağ, -90: yukarı, 90: aşağı)
  ctx.save();
  ctx.translate(x + s / 2, y + s / 2);
  ctx.rotate((angleDeg * Math.PI) / 180);
  ctx.fillStyle = color;
  const h = s / 2;
  ctx.beginPath();
  ctx.moveTo(-h * 0.85, -h * 0.22);
  ctx.lineTo(h * 0.1, -h * 0.22);
  ctx.lineTo(h * 0.1, -h * 0.7);
  ctx.lineTo(h * 0.9, 0);
  ctx.lineTo(h * 0.1, h * 0.7);
  ctx.lineTo(h * 0.1, h * 0.22);
  ctx.lineTo(-h * 0.85, h * 0.22);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

function circle(ctx: Ctx, x: number, y: number, s: number, color: string): void {
  ctx.fillStyle = color;
  ctx.beginPath();
  ctx.arc(x + s / 2, y + s / 2, s * 0.42, 0, Math.PI * 2);
  ctx.fill();
}

function triangle(ctx: Ctx, x: number, y: number, s: number, up: boolean, color: string): void {
  ctx.fillStyle = color;
  ctx.beginPath();
  if (up) {
    ctx.moveTo(x + s / 2, y + s * 0.12);
    ctx.lineTo(x + s * 0.92, y + s * 0.85);
    ctx.lineTo(x + s * 0.08, y + s * 0.85);
  } else {
    ctx.moveTo(x + s * 0.08, y + s * 0.15);
    ctx.lineTo(x + s * 0.92, y + s * 0.15);
    ctx.lineTo(x + s / 2, y + s * 0.88);
  }
  ctx.closePath();
  ctx.fill();
}

function diamond(ctx: Ctx, x: number, y: number, s: number, color: string): void {
  ctx.fillStyle = color;
  ctx.beginPath();
  ctx.moveTo(x + s / 2, y + s * 0.08);
  ctx.lineTo(x + s * 0.92, y + s / 2);
  ctx.lineTo(x + s / 2, y + s * 0.92);
  ctx.lineTo(x + s * 0.08, y + s / 2);
  ctx.closePath();
  ctx.fill();
}

function mark(ctx: Ctx, x: number, y: number, s: number, kind: "x" | "!" | "check", color: string, ring: boolean): void {
  if (ring) circle(ctx, x, y, s, color);
  ctx.strokeStyle = ring ? "#FFFFFF" : color;
  ctx.fillStyle = ring ? "#FFFFFF" : color;
  ctx.lineWidth = Math.max(1.2, s * (ring ? 0.13 : 0.18));
  ctx.lineCap = "round";
  const c = s / 2;
  const k = ring ? 0.2 : 0.3;
  ctx.beginPath();
  if (kind === "x") {
    ctx.moveTo(x + c - s * k, y + c - s * k);
    ctx.lineTo(x + c + s * k, y + c + s * k);
    ctx.moveTo(x + c + s * k, y + c - s * k);
    ctx.lineTo(x + c - s * k, y + c + s * k);
    ctx.stroke();
  } else if (kind === "check") {
    ctx.moveTo(x + c - s * k, y + c);
    ctx.lineTo(x + c - s * 0.05, y + c + s * k * 0.8);
    ctx.lineTo(x + c + s * k, y + c - s * k * 0.9);
    ctx.stroke();
  } else {
    ctx.moveTo(x + c, y + c - s * k);
    ctx.lineTo(x + c, y + c + s * 0.05);
    ctx.stroke();
    ctx.beginPath();
    ctx.arc(x + c, y + c + s * k * 0.85, ctx.lineWidth * 0.55, 0, Math.PI * 2);
    ctx.fill();
  }
}

function flag(ctx: Ctx, x: number, y: number, s: number, color: string): void {
  ctx.fillStyle = BLACK;
  ctx.fillRect(x + s * 0.22, y + s * 0.1, Math.max(1, s * 0.08), s * 0.8);
  ctx.fillStyle = color;
  ctx.beginPath();
  ctx.moveTo(x + s * 0.3, y + s * 0.1);
  ctx.lineTo(x + s * 0.88, y + s * 0.3);
  ctx.lineTo(x + s * 0.3, y + s * 0.52);
  ctx.closePath();
  ctx.fill();
}

function star(ctx: Ctx, x: number, y: number, s: number, fill: number): void {
  const path = () => {
    ctx.beginPath();
    for (let i = 0; i < 10; i++) {
      const r = i % 2 === 0 ? s * 0.46 : s * 0.19;
      const a = -Math.PI / 2 + (i * Math.PI) / 5;
      const px = x + s / 2 + r * Math.cos(a);
      const py = y + s / 2 + 0.03 * s + r * Math.sin(a);
      if (i === 0) ctx.moveTo(px, py);
      else ctx.lineTo(px, py);
    }
    ctx.closePath();
  };
  ctx.save();
  path();
  ctx.fillStyle = "#FFFFFF";
  ctx.fill();
  if (fill > 0) {
    ctx.save();
    ctx.clip();
    ctx.fillStyle = YELLOW;
    ctx.fillRect(x, y, s * fill, s);
    ctx.restore();
  }
  path();
  ctx.strokeStyle = "#C99A00";
  ctx.lineWidth = 1;
  ctx.stroke();
  ctx.restore();
}

function bars(ctx: Ctx, x: number, y: number, s: number, n: number, on: number): void {
  const w = (s * 0.8) / n;
  for (let i = 0; i < n; i++) {
    const h = (s * 0.8 * (i + 1)) / n;
    ctx.fillStyle = i < on ? "#2F6EBA" : "#C8C8C8";
    ctx.fillRect(x + s * 0.1 + i * w + w * 0.12, y + s * 0.9 - h, w * 0.76, h);
  }
}

function quarter(ctx: Ctx, x: number, y: number, s: number, q: number): void {
  const cx = x + s / 2;
  const cy = y + s / 2;
  const r = s * 0.42;
  ctx.fillStyle = "#FFFFFF";
  ctx.beginPath();
  ctx.arc(cx, cy, r, 0, Math.PI * 2);
  ctx.fill();
  if (q > 0) {
    ctx.fillStyle = BLACK;
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, r, -Math.PI / 2, -Math.PI / 2 + (Math.PI * 2 * q) / 4);
    ctx.closePath();
    ctx.fill();
  }
  ctx.strokeStyle = BLACK;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.arc(cx, cy, r, 0, Math.PI * 2);
  ctx.stroke();
}

function boxes(ctx: Ctx, x: number, y: number, s: number, on: number): void {
  const cell = s * 0.4;
  const pos = [
    [0, 0],
    [1, 0],
    [0, 1],
    [1, 1],
  ];
  pos.forEach(([i, j], k) => {
    ctx.fillStyle = k < on ? "#2F6EBA" : "#C8C8C8";
    ctx.fillRect(x + s * 0.08 + i * (cell + s * 0.04), y + s * 0.08 + j * (cell + s * 0.04), cell, cell);
  });
}

/** Simgeyi (küme, düşükten yükseğe dizin) verilen kareye çizer. */
export function drawIcon(ctx: Ctx, set: string, idx: number, x: number, y: number, s: number): void {
  ctx.save();
  switch (set) {
    case "3Arrows":
      arrow(ctx, x, y, s, [90, 0, -90][idx], [RED, YELLOW, GREEN][idx]);
      break;
    case "3ArrowsGray":
      arrow(ctx, x, y, s, [90, 0, -90][idx], GRAY);
      break;
    case "4Arrows":
      arrow(ctx, x, y, s, [90, 45, -45, -90][idx], [RED, YELLOW, YELLOW, GREEN][idx]);
      break;
    case "4ArrowsGray":
      arrow(ctx, x, y, s, [90, 45, -45, -90][idx], GRAY);
      break;
    case "5Arrows":
      arrow(ctx, x, y, s, [90, 45, 0, -45, -90][idx], [RED, YELLOW, YELLOW, YELLOW, GREEN][idx]);
      break;
    case "5ArrowsGray":
      arrow(ctx, x, y, s, [90, 45, 0, -45, -90][idx], GRAY);
      break;
    case "3Triangles":
      if (idx === 1) {
        ctx.fillStyle = YELLOW;
        ctx.fillRect(x + s * 0.2, y + s * 0.42, s * 0.6, s * 0.16);
      } else triangle(ctx, x, y, s, idx === 2, idx === 2 ? GREEN : RED);
      break;
    case "3TrafficLights1":
    case "3TrafficLights2":
      circle(ctx, x, y, s, [RED, YELLOW, GREEN][idx]);
      break;
    case "4TrafficLights":
      circle(ctx, x, y, s, [BLACK, RED, YELLOW, GREEN][idx]);
      break;
    case "4RedToBlack":
      circle(ctx, x, y, s, [BLACK, GRAY, "#F4A6A0", RED][idx]);
      break;
    case "3Signs":
      if (idx === 0) diamond(ctx, x, y, s, RED);
      else if (idx === 1) triangle(ctx, x, y, s, true, YELLOW);
      else circle(ctx, x, y, s, GREEN);
      break;
    case "3Symbols":
      mark(ctx, x, y, s, (["x", "!", "check"] as const)[idx], [RED, YELLOW, GREEN][idx], true);
      break;
    case "3Symbols2":
      mark(ctx, x, y, s, (["x", "!", "check"] as const)[idx], [RED, YELLOW, GREEN][idx], false);
      break;
    case "3Flags":
      flag(ctx, x, y, s, [RED, YELLOW, GREEN][idx]);
      break;
    case "3Stars":
      star(ctx, x, y, s, [0, 0.5, 1][idx]);
      break;
    case "4Rating":
      bars(ctx, x, y, s, 4, idx + 1);
      break;
    case "5Rating":
      bars(ctx, x, y, s, 4, idx);
      break;
    case "5Quarters":
      quarter(ctx, x, y, s, idx);
      break;
    case "5Boxes":
      boxes(ctx, x, y, s, idx);
      break;
    default:
      circle(ctx, x, y, s, [RED, YELLOW, GREEN][Math.min(2, idx)]);
  }
  ctx.restore();
}
