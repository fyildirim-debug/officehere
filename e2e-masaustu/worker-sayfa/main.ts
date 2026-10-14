// Masaüstü Worker duman testinin sayfası: Worker RPC'yi uygulamaların kullanacağı biçimde kurar
// (Vite modül Worker'ı, file://, uygulama CSP'si) ve durum çubuğunda ilerleme göstergesini çizer.
import { createElement, useSyncExternalStore } from "react";
import { createRoot } from "react-dom/client";
import { IoWorkerError, SaveError, createIoWorker, isAbortError } from "../../packages/core/src/index";
import { Progress, type ProgressProps } from "../../packages/ui/src/components/Progress";
import { StatusBar } from "../../packages/ui/src/components/StatusBar";
import type { handlers } from "./islemler";
import { FormatError } from "./ortak";

let made = 0;
const io = createIoWorker<typeof handlers>(
  () => {
    made++;
    return new Worker(new URL("./io.worker.ts", import.meta.url), { type: "module" });
  },
  { errors: { FormatError } },
);

/* Durum çubuğu */

let progress: ProgressProps | null = null;
const subs = new Set<() => void>();
function show(p: ProgressProps | null) {
  progress = p;
  for (const fn of subs) fn();
}
function App() {
  const p = useSyncExternalStore(
    (cb) => {
      subs.add(cb);
      return () => {
        subs.delete(cb);
      };
    },
    () => progress,
  );
  return createElement(StatusBar, { left: p ? createElement(Progress, p) : createElement("span", { className: "oh-status-item" }, "Hazır") });
}
createRoot(document.getElementById("root")!).render(createElement(App));

/* Karşıtlık ölçümü */

function rgb(color: string): [number, number, number] {
  const c = document.createElement("canvas");
  c.width = c.height = 1;
  const g = c.getContext("2d")!;
  g.fillStyle = color;
  g.fillRect(0, 0, 1, 1);
  const d = g.getImageData(0, 0, 1, 1).data;
  return [d[0], d[1], d[2]];
}
function luminance([r, g, b]: [number, number, number]) {
  const ch = (v: number) => {
    const s = v / 255;
    return s <= 0.04045 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b);
}
function contrast(a: string, b: string) {
  const [x, y] = [luminance(rgb(a)), luminance(rgb(b))].sort((p, q) => q - p);
  return Math.round(((x + 0.05) / (y + 0.05)) * 100) / 100;
}
const css = (sel: string) => getComputedStyle(document.querySelector(sel)!);

/* Test arayüzü */

type Run = { abortedAt: number; settledAt: number; result: string; error: string; progress: number[] };
let run: Run | null = null;

export const api = {
  kurulan: () => made,
  ortam: () => io.call("ortam", undefined),
  async ping() {
    const t0 = performance.now();
    const id = await io.call("ping", undefined);
    return { id, ms: Math.round(performance.now() - t0) };
  },
  async aktar(mb: number) {
    const bytes = new Uint8Array(mb << 20);
    let expected = 0;
    for (let i = 0; i < bytes.length; i += 997) {
      bytes[i] = i & 255;
      expected = (expected + (i & 255)) >>> 0;
    }
    const buf = bytes.buffer;
    const t0 = performance.now();
    const p = io.call("topla", { bytes: buf }, { transfer: [buf] });
    const detached = buf.byteLength === 0;
    const out = await p;
    return { detached, expected, sum: out.sum, back: out.bytes.byteLength, ms: Math.round(performance.now() - t0) };
  },
  async hata() {
    const f = (await io.call("bozuk", undefined).catch((e: unknown) => e)) as FormatError;
    const s = (await io.call("dolu", undefined).catch((e: unknown) => e)) as SaveError;
    return { format: f instanceof FormatError, name: f.name, message: f.message, detail: f.detail, save: s instanceof SaveError, code: s.code, reason: s.reason };
  },
  /** Uzun eşzamanlı işlem; durum çubuğundaki İptal düğmesi AbortController'ı tetikler. */
  baslat(ms: number) {
    const ctrl = new AbortController();
    const r: Run = { abortedAt: 0, settledAt: 0, result: "", error: "", progress: [] };
    run = r;
    const cancel = () => {
      r.abortedAt = performance.now();
      ctrl.abort();
    };
    show({ label: "Açılıyor…", value: 0, onCancel: cancel });
    io.call("mesgul", { ms }, { signal: ctrl.signal, onProgress: (p) => (r.progress.push(p.ratio), show({ label: "Açılıyor…", value: p.ratio, detail: p.step, onCancel: cancel })) })
      .then(
        (v) => (r.result = v),
        (e: unknown) => (r.error = isAbortError(e) ? "AbortError" : String(e)),
      )
      .finally(() => {
        r.settledAt = performance.now();
        show(null);
      });
  },
  sonuc: () => run && { ...run, progress: run.progress.length, cancelMs: run.settledAt && run.abortedAt ? Math.round((run.settledAt - run.abortedAt) * 10) / 10 : null },
  goster: (value: number | null) => show({ label: "Açılıyor…", value, onCancel: () => undefined }),
  karsitlik(app: string, theme: string) {
    document.documentElement.dataset.app = app;
    document.documentElement.dataset.ohTheme = theme;
    const bar = css(".oh-statusbar").backgroundColor;
    const fill = css(".oh-progress-fill").backgroundColor;
    return {
      fill: rgb(fill).join(","),
      dolguIz: contrast(fill, css(".oh-progress-track").backgroundColor),
      dolguZemin: contrast(fill, bar),
      metin: contrast(css(".oh-progress-label").color, bar),
      dugme: contrast(css(".oh-progress-cancel").color, bar),
    };
  },
  async yuklenemez() {
    const bad = createIoWorker<typeof handlers>(() => new Worker(new URL("./yok.js", location.href), { type: "module" }));
    const e = (await bad.call("ping", undefined).catch((x: unknown) => x)) as IoWorkerError;
    return { cls: e instanceof IoWorkerError, code: e.code, message: e.message };
  },
};

(window as unknown as { ohWorkerTest: typeof api }).ohWorkerTest = api;
