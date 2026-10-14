// Tarayıcı performansı: soğuk açılış, büyük dosya açma, yazma gecikmesi, kaydırma FPS'i, küçük resimler, bellek.
// Uygulamaların dist/ derlemeleri yerel HTTP sunucusundan verilir; masaüstü köprüsü (window.officehereNative)
// sahte olarak enjekte edilir ve dosyalar `?file=` ile (Electron'daki gibi) açılır.
// Ön koşul: pnpm build (apps/*/dist), python3 scripts/derlem/olcum/perf_veri.py, Playwright Chromium.
// Kullanım: node scripts/derlem/olcum/perf_tarayici.mjs [senaryo...]
//   senaryolar: soğuk_belge soğuk_tablo soğuk_sunum belge300 tablo200k sunum200 (boşsa hepsi)
// Çıktı: .derlem/rapor/perf_tarayici[-senaryolar].json
import { chromium } from "playwright";
import { createServer } from "node:http";
import { existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { extname, join, normalize } from "node:path";
import { CIKTI, KOK, RAPOR } from "../ortak.mjs";

const VERI = join(CIKTI, "perf/veri");
const MIME = { ".html": "text/html", ".js": "text/javascript", ".mjs": "text/javascript", ".css": "text/css", ".svg": "image/svg+xml", ".png": "image/png", ".woff2": "font/woff2", ".json": "application/json", ".txt": "text/plain", ".wasm": "application/wasm" };
const server = createServer((req, res) => {
  const u = new URL(req.url, "http://x");
  let p = decodeURIComponent(u.pathname);
  let file;
  if (p.startsWith("/__veri/")) file = join(VERI, p.slice(8));
  else {
    const m = p.match(/^\/(belge|tablo|sunum)(\/.*)?$/);
    if (!m) return res.writeHead(404).end();
    const rel = !m[2] || m[2] === "/" ? "/index.html" : m[2];
    file = normalize(join(KOK, "apps", m[1], "dist", rel));
  }
  if (!existsSync(file) || statSync(file).isDirectory()) return res.writeHead(404).end();
  res.writeHead(200, { "content-type": MIME[extname(file)] ?? "application/octet-stream", "cache-control": "no-store" });
  res.end(readFileSync(file));
});
await new Promise((r) => server.listen(0, r));
const PORT = server.address().port;
const BASE = `http://127.0.0.1:${PORT}`;

// Sahte masaüstü köprüsü + uzun görev / olay zamanlaması gözlemcileri
const INIT = `
(() => {
  window.__lt = []; window.__ev = []; window.__writes = [];
  try { new PerformanceObserver((l) => { for (const e of l.getEntries()) window.__lt.push([Math.round(e.startTime), Math.round(e.duration)]); }).observe({ type: "longtask", buffered: true }); } catch {}
  try { new PerformanceObserver((l) => { for (const e of l.getEntries()) window.__ev.push([e.name, Math.round(e.startTime), Math.round(e.duration), Math.round(e.processingEnd - e.processingStart)]); }).observe({ type: "event", durationThreshold: 16, buffered: true }); } catch {}
  if (!location.search.includes("native=0")) window.officehereNative = {
    platform: "linux",
    openFile: async () => null,
    readFile: async (p) => { const name = p.split("/").pop(); const r = await fetch("/__veri/" + encodeURIComponent(name)); if (!r.ok) return null; return { path: p, name, data: await r.arrayBuffer() }; },
    saveFileAs: async () => null,
    writeFile: async (p, d) => { window.__writes.push([Math.round(performance.now()), p.split("/").pop(), d.byteLength]); return true; },
    getRecent: async () => [], addRecent: async () => {}, pinRecent: async () => {}, removeRecent: async () => {}, showInFolder: async () => {},
    recentStatus: async () => ({}), getSettings: async () => ({}), setSetting: async () => {}, newWindow: async () => {},
    setDocumentState: () => {}, onOpenPath: () => () => {}, onBeforeClose: () => () => {}, setTitleBarOverlay: () => {}, exportPdf: async () => null,
    updates: { getState: async () => null, check: async () => null, install: async () => null, setAutoCheck: async () => {}, onChange: () => () => {} },
  };
})();`;

const browser = await chromium.launch({ args: ["--enable-precise-memory-info", "--disable-gpu-vsync=false"] });
const results = {};
const med = (a) => { const s = [...a].sort((x, y) => x - y); return s.length ? s[Math.floor(s.length / 2)] : null; };
const pct = (a, p) => { const s = [...a].sort((x, y) => x - y); return s.length ? s[Math.min(s.length - 1, Math.floor(s.length * p))] : null; };

async function newPage() {
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 880 }, locale: "tr-TR", timezoneId: "Europe/Istanbul" });
  await ctx.addInitScript(INIT);
  const page = await ctx.newPage();
  const cdp = await ctx.newCDPSession(page);
  await cdp.send("Performance.enable");
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => m.type() === "error" && errors.push("console: " + m.text().slice(0, 200)));
  return { ctx, page, cdp, errors };
}
async function heapMB(cdp) {
  await cdp.send("HeapProfiler.collectGarbage").catch(() => {});
  const h = await cdp.send("Runtime.getHeapUsage");
  return Math.round(h.usedSize / 1048576);
}
async function waitFor(page, fn, arg, timeout = 180000, poll = 100) {
  const t0 = Date.now();
  while (Date.now() - t0 < timeout) {
    if (await page.evaluate(fn, arg).catch(() => false)) return true;
    await page.waitForTimeout(poll);
  }
  return false;
}
const titleHas = (s) => !!document.querySelector(".oh-titlebar")?.textContent?.includes(s);

/* ---------------------------------------------------------------- 4. Soğuk açılış */
async function coldStart(app, ready) {
  const runs = [];
  for (let i = 0; i < 5; i++) {
    const { ctx, page, cdp, errors } = await newPage();
    await cdp.send("Network.enable");
    await cdp.send("Network.setCacheDisabled", { cacheDisabled: true });
    const t0 = Date.now();
    await page.goto(`${BASE}/${app}/`);
    await page.waitForSelector(ready, { timeout: 60000 });
    const wall = Date.now() - t0;
    await page.waitForTimeout(300);
    const m = await page.evaluate(() => {
      const nav = performance.getEntriesByType("navigation")[0];
      const fcp = performance.getEntriesByName("first-contentful-paint")[0];
      const res = performance.getEntriesByType("resource");
      return { dcl: Math.round(nav.domContentLoadedEventEnd), load: Math.round(nav.loadEventEnd), fcp: fcp ? Math.round(fcp.startTime) : null, js: res.filter((r) => r.initiatorType === "script" || r.name.endsWith(".js")).length, bytes: res.reduce((a, r) => a + (r.encodedBodySize || 0), 0), lt: window.__lt.reduce((a, x) => a + x[1], 0) };
    });
    runs.push({ ...m, hazır: wall, heapMB: await heapMB(cdp), hata: errors.length });
    await ctx.close();
  }
  const pick = (k) => med(runs.map((r) => r[k]));
  return { hazırMs: pick("hazır"), fcpMs: pick("fcp"), dclMs: pick("dcl"), loadMs: pick("load"), jsİstek: pick("js"), aktarılanKB: Math.round(pick("bytes") / 1024), uzunGörevToplamMs: pick("lt"), heapMB: pick("heapMB"), hata: runs.reduce((a, r) => a + r.hata, 0), tümü: runs.map((r) => r.hazır) };
}

/* ---------------------------------------------------------------- 3a. Belge: 300 sayfa */
async function belgeBig() {
  const { ctx, page, cdp, errors } = await newPage();
  const out = {};
  await page.goto(`${BASE}/belge/`);
  await page.waitForSelector(".ProseMirror");
  out.boşHeapMB = await heapMB(cdp);
  const t0 = Date.now();
  const BD = process.env.BELGE_DOSYA ?? "buyuk-belge.docx";
  await page.goto(`${BASE}/belge/?file=/veri/${BD}`);
  const ok = BD === "buyuk-belge.docx" ? await waitFor(page, () => document.querySelector(".ProseMirror")?.textContent?.includes("SON-PARAGRAF-İŞARETİ"), null, 240000) : await waitFor(page, titleHas, BD.replace(/\.docx$/, ""), 60000);
  out.metinYüklendiMs = ok ? Date.now() - t0 : "ZAMAN AŞIMI";
  // sayfalama: .bl-page sayısı 2 sn boyunca değişmeyene dek
  let last = -1, stableSince = Date.now(), pages = 0;
  while (Date.now() - stableSince < 2000 && Date.now() - t0 < 300000) {
    pages = await page.evaluate(() => document.querySelectorAll(".bl-page").length);
    if (pages !== last) { last = pages; stableSince = Date.now(); }
    await page.waitForTimeout(100);
  }
  out.sayfalamaBittiMs = stableSince - t0;
  out.sayfa = pages;
  out.durumÇubuğu = await page.evaluate(() => [...document.querySelectorAll(".oh-statusbar *")].map((e) => e.textContent).find((t) => /Sayfa \d/.test(t ?? "")) ?? "");
  out.domDüğüm = await page.evaluate(() => document.getElementsByTagName("*").length);
  out.açılışUzunGörev = await page.evaluate(() => ({ adet: window.__lt.length, toplamMs: window.__lt.reduce((a, x) => a + x[1], 0), enUzunMs: Math.max(0, ...window.__lt.map((x) => x[1])) }));
  out.açıkHeapMB = await heapMB(cdp);
  // yazma gecikmesi: belge başına tıkla, 60 karakter yaz
  const typing = async (label, place) => {
    await place();
    await page.evaluate(() => {
      window.__keys = [];
      window.__evStart = window.__ev.length;
      window.__ltStart = window.__lt.length;
      window.__kd = (e) => {
        const t = performance.now();
        requestAnimationFrame(() => { const t1 = performance.now(); setTimeout(() => window.__keys.push([Math.round(t1 - t), Math.round(performance.now() - t)]), 0); });
      };
      document.addEventListener("keydown", window.__kd, true);
    });
    await page.keyboard.type("Performans ölçüm metni yazılıyor ığüşöç 0123456789 abcdefghij", { delay: 40 });
    await page.waitForTimeout(500);
    const r = await page.evaluate(() => {
      document.removeEventListener("keydown", window.__kd, true);
      const ev = window.__ev.slice(window.__evStart).filter((e) => /key|input|beforeinput/.test(e[0]));
      return { keys: window.__keys, ev, lt: window.__lt.slice(window.__ltStart) };
    });
    const lat = r.keys.map((k) => k[1]);
    return { [label]: { tuş: lat.length, medyanMs: med(lat), p95Ms: pct(lat, 0.95), enKötüMs: Math.max(...lat), olay16msÜstü: r.ev.length, olayEnUzunMs: Math.max(0, ...r.ev.map((e) => e[2])), uzunGörev: r.lt.length } };
  };
  Object.assign(out, await typing("yazmaBaşta", async () => { await page.click(".ProseMirror p >> nth=1"); await page.keyboard.press("End"); }));
  Object.assign(out, await typing("yazmaSonda", async () => { await page.keyboard.press("Control+End"); await page.waitForTimeout(1500); }));
  // otomatik kaydetme / taslak maliyeti: 6 sn boşta bekle, uzun görevleri ve yazmaları say
  const before = await page.evaluate(() => [window.__lt.length, window.__writes.length]);
  await page.waitForTimeout(7000);
  out.boştaOtomatikKaydet = await page.evaluate((b) => ({ uzunGörev: window.__lt.slice(b[0]).map((x) => x[1]), diskYazma: window.__writes.slice(b[1]).map((w) => `${w[1]} ${Math.round(w[2] / 1024)}KB`) }), before);
  out.taslakKB = await page.evaluate(() => { let n = 0; for (let i = 0; i < localStorage.length; i++) { const k = localStorage.key(i); if (k.includes("draft")) n += (localStorage.getItem(k) ?? "").length; } return Math.round(n / 1024); });
  out.sonHeapMB = await heapMB(cdp);
  out.hata = errors.slice(0, 5);
  await ctx.close();
  return out;
}

/* ---------------------------------------------------------------- 3b. Tablo: 200k satır + 10k zincir */
async function frameStats(page, ms, driver) {
  await page.evaluate(() => {
    window.__frames = [];
    window.__fon = true;
    const loop = (t) => { if (!window.__fon) return; window.__frames.push(t); requestAnimationFrame(loop); };
    requestAnimationFrame(loop);
  });
  const t0 = Date.now();
  while (Date.now() - t0 < ms) await driver();
  const f = await page.evaluate(() => { window.__fon = false; return window.__frames; });
  const d = f.slice(1).map((t, i) => t - f[i]);
  const dur = (f[f.length - 1] - f[0]) / 1000;
  return { fps: Math.round((f.length - 1) / dur), kareMedyanMs: Math.round(med(d)), kareP95Ms: Math.round(pct(d, 0.95)), enUzunKareMs: Math.round(Math.max(...d)), uzun50msÜstü: d.filter((x) => x > 50).length };
}
async function tabloBig() {
  const { ctx, page, cdp, errors } = await newPage();
  const out = {};
  await page.goto(`${BASE}/tablo/`);
  await page.waitForSelector(".tb-canvas");
  out.boşHeapMB = await heapMB(cdp);
  const t0 = Date.now();
  await page.goto(`${BASE}/tablo/?file=/veri/buyuk-tablo.xlsx`);
  const ok = await waitFor(page, titleHas, "buyuk-tablo", 300000);
  out.açılışMs = ok ? Date.now() - t0 : "ZAMAN AŞIMI";
  await page.waitForTimeout(1000);
  out.açılışUzunGörev = await page.evaluate(() => ({ adet: window.__lt.length, toplamMs: window.__lt.reduce((a, x) => a + x[1], 0), enUzunMs: Math.max(0, ...window.__lt.map((x) => x[1])) }));
  out.açıkHeapMB = await heapMB(cdp);
  const box = await page.locator(".tb-canvas").boundingBox();
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  out.kaydırmaTekerlek = await frameStats(page, 4000, async () => { await page.mouse.wheel(0, 600); await page.waitForTimeout(16); });
  out.kaydırmaYatay = await frameStats(page, 2000, async () => { await page.mouse.wheel(600, 0); await page.waitForTimeout(16); });
  // Ctrl+End / Ctrl+Home süresi (olay zamanlaması)
  await page.mouse.click(box.x + 120, box.y + 60);
  const evStart = await page.evaluate(() => window.__ev.length);
  await page.keyboard.press("Control+End");
  await page.waitForTimeout(300);
  await page.keyboard.press("Control+Home");
  await page.waitForTimeout(300);
  out.ctrlEndHomeOlayMs = await page.evaluate((s) => window.__ev.slice(s).filter((e) => e[0] === "keydown").map((e) => e[2]), evStart);
  // bir hücreye yazma (200k satırlık kitapta)
  const evs = await page.evaluate(() => window.__ev.length);
  await page.keyboard.type("12345", { delay: 40 });
  await page.keyboard.press("Enter");
  await page.waitForTimeout(500);
  out.hücreYazmaOlayMs = await page.evaluate((s) => window.__ev.slice(s).map((e) => `${e[0]}:${e[2]}`), evs);
  const before = await page.evaluate(() => [window.__lt.length, window.__writes.length]);
  await page.waitForTimeout(45000);
  out.boştaOtomatikKaydet = await page.evaluate((b) => ({ uzunGörev: window.__lt.slice(b[0]).map((x) => x[1]), diskYazma: window.__writes.slice(b[1]).map((w) => `${w[1]} ${Math.round(w[2] / 1024)}KB`) }), before);
  out.sonHeapMB = await heapMB(cdp);
  out.hata = errors.slice(0, 5);
  await ctx.close();

  // 10k formül zinciri: A1 değiştir → zincir yeniden hesaplanır
  const z = await newPage();
  await z.page.goto(`${BASE}/tablo/?file=/veri/zincir-10k.xlsx`);
  await waitFor(z.page, titleHas, "zincir-10k", 60000);
  await z.page.waitForTimeout(800);
  const zb = await z.page.locator(".tb-canvas").boundingBox();
  const zi = await z.page.evaluate(() => window.__ev.length);
  const zr = [];
  for (const v of ["5", "7", "9"]) {
    await z.page.keyboard.press("Control+Home");
    const s = await z.page.evaluate(() => window.__ev.length);
    await z.page.keyboard.type(v);
    const t = Date.now();
    await z.page.keyboard.press("Enter");
    await z.page.waitForTimeout(600);
    zr.push(await z.page.evaluate((s) => window.__ev.slice(s).filter((e) => e[0] === "keydown").map((e) => e[2]), s));
  }
  out.zincir10k_EnterOlayMs = zr;
  out.zincirHata = z.errors.slice(0, 3);
  await z.ctx.close();
  return out;
}

/* ---------------------------------------------------------------- 3c. Sunum: 200 slayt */
async function sunumBig() {
  const { ctx, page, cdp, errors } = await newPage();
  const out = {};
  await page.goto(`${BASE}/sunum/`);
  await page.waitForSelector(".thumb-pane");
  out.boşHeapMB = await heapMB(cdp);
  const t0 = Date.now();
  await page.goto(`${BASE}/sunum/?file=/veri/buyuk-sunu.pptx`);
  const ok = await waitFor(page, titleHas, "buyuk-sunu", 120000);
  out.açılışMs = ok ? Date.now() - t0 : "ZAMAN AŞIMI";
  const okT = await waitFor(page, () => document.querySelectorAll(".thumb-item").length >= 200, null, 120000);
  out.küçükResim200Ms = okT ? Date.now() - t0 : "ZAMAN AŞIMI";
  out.küçükResimÇizilmiş = await page.evaluate(() => [...document.querySelectorAll(".thumb-item .thumb-box")].filter((b) => b.querySelector("svg, canvas, img, div")).length);
  out.domDüğüm = await page.evaluate(() => document.getElementsByTagName("*").length);
  out.açılışUzunGörev = await page.evaluate(() => ({ adet: window.__lt.length, toplamMs: window.__lt.reduce((a, x) => a + x[1], 0), enUzunMs: Math.max(0, ...window.__lt.map((x) => x[1])) }));
  out.açıkHeapMB = await heapMB(cdp);
  const pane = await page.locator(".thumb-pane").first().boundingBox();
  await page.mouse.move(pane.x + pane.width / 2, pane.y + pane.height / 2);
  out.küçükResimKaydırma = await frameStats(page, 3000, async () => { await page.mouse.wheel(0, 500); await page.waitForTimeout(16); });
  // slayt geçişi: PageDown x 20
  await page.mouse.click(pane.x + pane.width / 2, pane.y + 40);
  const s = await page.evaluate(() => window.__ev.length);
  for (let i = 0; i < 20; i++) { await page.keyboard.press("PageDown"); await page.waitForTimeout(60); }
  await page.waitForTimeout(400);
  const evd = await page.evaluate((s) => window.__ev.slice(s).filter((e) => e[0] === "keydown").map((e) => e[2]), s);
  out.slaytDeğiştirOlayMs = { adet16msÜstü: evd.length, medyan: med(evd), enKötü: Math.max(0, ...evd) };
  out.sonHeapMB = await heapMB(cdp);
  out.hata = errors.slice(0, 5);
  await ctx.close();
  return out;
}

const want = process.argv.slice(2);
const run = async (k, f) => {
  if (want.length && !want.includes(k)) return;
  const t = Date.now();
  try {
    results[k] = await f();
  } catch (e) {
    results[k] = { HATA: String(e.message).slice(0, 400) };
  }
  console.log(k, Math.round((Date.now() - t) / 1000) + "s", JSON.stringify(results[k]));
};
await run("soğuk_belge", () => coldStart("belge", ".ProseMirror"));
await run("soğuk_tablo", () => coldStart("tablo", ".tb-canvas"));
await run("soğuk_sunum", () => coldStart("sunum", ".thumb-pane"));
await run("belge300", belgeBig);
await run("tablo200k", tabloBig);
await run("sunum200", sunumBig);
mkdirSync(RAPOR, { recursive: true });
writeFileSync(join(RAPOR, `perf_tarayici${want.length ? "-" + want.join("-") : ""}.json`), JSON.stringify(results, null, 1));
await browser.close();
server.close();
