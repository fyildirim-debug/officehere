// Sayfadaki yakalanmamış hatalar: window "error" ve "unhandledrejection".
// Her hata (belge içeriği olmadan) yerel günlüğe yazılır (masaüstünde ana süreç,
// tarayıcıda bellekteki halka) ve kurtarma taslağı hemen yazdırılır
// ("officehere:autosave" olayı; uygulamaların taslak yazıcıları dinler).
// Hiçbir şey dışarı gönderilmez; "Hata bildir" penceresi son kayıtları kullanıcıya gösterir.
import { getNative, type NativeLogEntry } from "@officehere/core";

/** Taslak yazıcıları "şimdi yaz" diye uyaran olay. */
export const FLUSH_DRAFT_EVENT = "officehere:autosave";

const MAX_LINES = 60;
const ring: string[] = [];

/** Tarayıcıda (ve masaüstünde ek olarak) bu sayfanın son günlük satırları. */
export function recentPageLog(): string[] {
  return ring.slice();
}

function trimStack(stack: string | undefined): string | undefined {
  if (!stack) return undefined;
  // Yalnız kod konumları: "at f (index-abc.js:12:3)".
  const lines = stack
    .split("\n")
    .slice(1)
    .map((l) => l.trim())
    .filter((l) => l.startsWith("at "))
    .slice(0, 8)
    .map((l) => l.replace(/(?:file|https?):\/\/[^\s)]*\/([^/\s)]+)/g, "$1"));
  return lines.length ? `Error\n    ${lines.join("\n    ")}` : undefined;
}

/** Günlüğe bir kayıt yazar (masaüstünde dosyaya, her yerde bellekteki halkaya). */
export function logPageEvent(entry: NativeLogEntry): void {
  const message = String(entry.message ?? "").replace(/\s+/g, " ").slice(0, 400);
  const level = entry.level === "bilgi" ? "BİLGİ" : entry.level === "uyari" ? "UYARI" : "HATA";
  ring.push(`${new Date().toISOString()} ${level} [sayfa] ${message}${entry.kind ? ` (${entry.kind})` : ""}`);
  if (ring.length > MAX_LINES) ring.splice(0, ring.length - MAX_LINES);
  try {
    getNative()?.log?.({ ...entry, message, stack: trimStack(entry.stack) });
  } catch {
    // günlük yazılamazsa sayfa etkilenmez
  }
}

/** Zararsız tarayıcı uyarıları (hata sayılmaz). */
function benign(message: string): boolean {
  return /ResizeObserver loop (?:completed|limit)/i.test(message);
}

let lastFlush = 0;

function flushDrafts() {
  // Hata fırtınasında taslak yazıcıyı boğmamak için en çok 3 sn'de bir.
  const now = Date.now();
  if (now - lastFlush < 3000) return;
  lastFlush = now;
  try {
    window.dispatchEvent(new Event(FLUSH_DRAFT_EVENT));
  } catch {
    // yok say
  }
}

function describeReason(reason: unknown): { message: string; stack?: string } {
  if (reason instanceof Error) return { message: `${reason.name}: ${reason.message}`, stack: reason.stack };
  if (typeof reason === "string") return { message: reason };
  try {
    return { message: JSON.stringify(reason)?.slice(0, 200) ?? String(reason) };
  } catch {
    return { message: String(reason) };
  }
}

let installed = false;

/** Yakalanmamış hata dinleyicilerini bir kez kurar (her uygulama ve başlangıç ekranı). */
export function installErrorReporting(): void {
  if (installed || typeof window === "undefined" || typeof window.addEventListener !== "function") return;
  installed = true;
  window.addEventListener("error", (e: ErrorEvent) => {
    // Kaynak (resim/betik) yükleme hataları ErrorEvent değildir.
    if (!(e instanceof ErrorEvent)) return;
    const message = e.error instanceof Error ? `${e.error.name}: ${e.error.message}` : e.message || "Bilinmeyen hata";
    if (benign(message)) return;
    logPageEvent({ level: "hata", message, stack: e.error instanceof Error ? e.error.stack : undefined, kind: "window.error" });
    flushDrafts();
  });
  window.addEventListener("unhandledrejection", (e: PromiseRejectionEvent) => {
    const { message, stack } = describeReason(e.reason);
    if (benign(message)) return;
    logPageEvent({ level: "hata", message, stack, kind: "unhandledrejection" });
    flushDrafts();
  });
}

installErrorReporting();
