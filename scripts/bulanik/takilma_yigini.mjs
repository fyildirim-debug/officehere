// Takılan (sonsuz döngü / çok uzun) bir örnekte yığın izini alır: alt süreci --inspect ile başlatır,
// N sn sonra Debugger.pause ile durdurup çağrı çerçevelerini kaynak haritasıyla özgün dosya:satıra çevirir.
// Önce paketler: node scripts/derlem/paketle.mjs (pnpm bulanik bunu yapar).
// Kullanım: node scripts/bulanik/takilma_yigini.mjs <belge|tablo|sunum> <dosya> [bekleme_sn=8]
//   ör. node scripts/bulanik/takilma_yigini.mjs tablo .derlem/bulanik/tablo/123_lo-butce_birlesik_tum.xlsx
import { spawn } from "node:child_process";
import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { SourceMap } from "node:module";
import { basename, join, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { CIKTI, KOK, PAKET } from "../derlem/ortak.mjs";

const [app, dosyaArg, beklemeSn = "8"] = process.argv.slice(2);
if (!["belge", "tablo", "sunum"].includes(app) || !dosyaArg) {
  console.error("Kullanım: node scripts/bulanik/takilma_yigini.mjs <belge|tablo|sunum> <dosya> [bekleme_sn]");
  process.exit(2);
}
const dosya = resolve(dosyaArg);
const ad = basename(dosya).replace(/^\d+_/, "");
mkdirSync(CIKTI, { recursive: true });
const kosucu = join(CIKTI, `_takilma_${process.pid}.mjs`);
const paket = pathToFileURL(join(PAKET, `${app}-io.mjs`)).href;
writeFileSync(
  kosucu,
  `import ${JSON.stringify(pathToFileURL(join(KOK, "scripts/derlem/dom.mjs")).href)};
import { readFileSync } from "node:fs";
const M = await import(${JSON.stringify(paket)});
const b = readFileSync(${JSON.stringify(dosya)});
const ab = b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
console.log("BAŞLA");
${app === "belge" ? `const r = await M.openDocument(${JSON.stringify(ad)}, ab);` : app === "tablo" ? `const r = await M.openBook(${JSON.stringify(ad)}, ab);` : `const r = await M.readPresentationFile(new Uint8Array(ab), ${JSON.stringify(ad)});`}
console.log("AÇILDI");
${app === "tablo" ? "await M.writeXlsx(r.wb);" : app === "belge" ? "await M.writeDocx(r.doc);" : 'await M.writePresentationFormat(r.pres, "pptx");'}
console.log("BİTTİ");
process.exit(0);
`,
);
const port = 9400 + (process.pid % 500);
const cocuk = spawn(process.execPath, [`--inspect=127.0.0.1:${port}`, kosucu], { stdio: ["ignore", "pipe", "pipe"] });
cocuk.stdout.on("data", (d) => {
  if (String(d).includes("AÇILDI")) console.log("  (açma tamamlandı; kaydetme sürüyor)");
  if (String(d).includes("BİTTİ")) console.log("bitti (takılmadı)");
});
let bitti = false;
cocuk.on("exit", () => (bitti = true));
const temizle = () => rmSync(kosucu, { force: true });
await new Promise((r) => setTimeout(r, 1500));
let ws;
try {
  if (bitti) throw new Error("süreç bitti");
  const liste = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
  ws = new WebSocket(liste[0].webSocketDebuggerUrl);
  await new Promise((r, j) => {
    ws.addEventListener("open", r);
    ws.addEventListener("error", j);
    setTimeout(() => j(new Error("hata ayıklayıcıya bağlanılamadı")), 5000);
  });
} catch (e) {
  console.log(bitti ? "süreç beklemeden bitti (takılma yok)" : `bağlanılamadı: ${e.message}`);
  cocuk.kill("SIGKILL");
  temizle();
  process.exit(0);
}
let id = 0;
const bekleyen = new Map();
const duraklatildi = new Promise((resolve) => {
  ws.addEventListener("message", (ev) => {
    const m = JSON.parse(ev.data);
    if (m.id && bekleyen.has(m.id)) bekleyen.get(m.id)(m);
    if (m.method === "Debugger.paused") resolve(m.params);
  });
});
const gonder = (method, params = {}) =>
  new Promise((r) => {
    const i = ++id;
    bekleyen.set(i, r);
    ws.send(JSON.stringify({ id: i, method, params }));
  });
await gonder("Debugger.enable");
await gonder("Runtime.runIfWaitingForDebugger");
await new Promise((r) => setTimeout(r, Number(beklemeSn) * 1000));
await gonder("Debugger.pause");
const p = await Promise.race([duraklatildi, new Promise((r) => setTimeout(() => r(null), 30000))]);
const harita = new SourceMap(JSON.parse(readFileSync(join(PAKET, `${app}-io.mjs.map`), "utf8")));
if (!p) console.log("duraklatılamadı (süreç bitmiş olabilir)");
else {
  console.log(`${dosya} — ${beklemeSn} sn sonra yığın:`);
  for (const f of p.callFrames.slice(0, 14)) {
    const { lineNumber, columnNumber } = f.location;
    let kaynak = `${f.url.split("/").pop()}:${lineNumber + 1}`;
    if (!f.url || f.url.includes(`${app}-io.mjs`)) {
      const e = harita.findEntry(lineNumber, columnNumber);
      if (e?.originalSource) kaynak = `${e.originalSource.replace(/^.*?(apps|packages|node_modules)\//, "$1/")}:${e.originalLine + 1}`;
    }
    console.log(`  at ${f.functionName || "(anonim)"} ${kaynak}`);
  }
}
cocuk.kill("SIGKILL");
ws.close();
temizle();
process.exit(0);
