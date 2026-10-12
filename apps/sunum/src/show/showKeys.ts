/** Slayt gösterisi klavye kısayolları (ilerle/geri, numara+Enter, siyah/beyaz ekran, ızgara, çıkış). */

export interface ShowKeyCtx {
  advance: () => void;
  back: () => void;
  first: () => void;
  last: () => void;
  /** 1 tabanlı slayt numarasına git. */
  gotoNumber: (n: number) => void;
  screen: "none" | "black" | "white";
  setScreen: (s: "none" | "black" | "white") => void;
  exit: () => void;
  grid: boolean;
  setGrid: (v: boolean) => void;
  numBuf: string;
  setNumBuf: (fn: (b: string) => string) => void;
  /** Bilgi noktası: yalnız Esc çalışır. */
  kiosk: boolean;
  /** Önizlemede ızgara yok. */
  preview: boolean;
}

export function handleShowKey(e: React.KeyboardEvent, c: ShowKeyCtx): void {
  const k = e.key;
  if (c.grid && k === "Escape") {
    c.setGrid(false);
    return;
  }
  if (c.kiosk) {
    if (k === "Escape") {
      e.preventDefault();
      c.exit();
    }
    return;
  }
  if (/^[0-9]$/.test(k)) {
    c.setNumBuf((b) => (b + k).slice(-4));
    return;
  }
  if (k === "Enter" && c.numBuf) {
    e.preventDefault();
    const n = parseInt(c.numBuf, 10);
    c.setNumBuf(() => "");
    c.gotoNumber(n);
    return;
  }
  c.setNumBuf(() => "");
  if (c.screen !== "none" && k !== "Escape") {
    c.setScreen("none");
    if (k === "b" || k === "B" || k === "w" || k === "W" || k === "." || k === ",") return;
  }
  switch (k) {
    case "ArrowRight":
    case "ArrowDown":
    case " ":
    case "PageDown":
    case "Enter":
    case "n":
    case "N":
      e.preventDefault();
      c.advance();
      break;
    case "ArrowLeft":
    case "ArrowUp":
    case "PageUp":
    case "Backspace":
    case "p":
    case "P":
      e.preventDefault();
      c.back();
      break;
    case "Home":
      c.first();
      break;
    case "End":
      c.last();
      break;
    case "b":
    case "B":
    case ".":
      c.setScreen("black");
      break;
    case "w":
    case "W":
    case ",":
      c.setScreen("white");
      break;
    // G ya da "-": tüm slaytları gör (slayt ızgarası).
    case "g":
    case "G":
    case "-":
      e.preventDefault();
      if (!c.preview) c.setGrid(true);
      break;
    case "Escape":
      e.preventDefault();
      if (c.screen !== "none") c.setScreen("none");
      else c.exit();
      break;
    default:
      break;
  }
}

/** "0:00:05" biçiminde süre. */
export function formatTime(ms: number, withHours = false): string {
  const s = Math.max(0, Math.floor(ms / 1000));
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const ss = s % 60;
  return `${h > 0 || withHours ? `${h}:` : ""}${String(m).padStart(2, "0")}:${String(ss).padStart(2, "0")}`;
}
