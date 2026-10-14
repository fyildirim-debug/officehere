// Performans ölçümü (model/G-Ç düzeyi): büyük dosyaları uygulama okuyucu/yazıcılarıyla açma-yazma süresi
// ve model belleği. Önce büyük dosyalar: python3 scripts/derlem/olcum/perf_veri.py
// Kullanım: node --expose-gc --max-old-space-size=8192 scripts/derlem/olcum/perf_node.mjs [--paketleme-yok]
// Çıktı: .derlem/rapor/perf_node.json
import "../dom.mjs";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { pathToFileURL } from "node:url";
import { argumanlar, CIKTI, PAKET, RAPOR } from "../ortak.mjs";
import { paketle } from "../paketle.mjs";

if (!argumanlar()["paketleme-yok"]) await paketle();
const yukle = (app) => import(pathToFileURL(join(PAKET, `${app}-io.mjs`)).href);
const [B, T, S] = await Promise.all([yukle("belge"), yukle("tablo"), yukle("sunum")]);
const V = join(CIKTI, "perf/veri");
const ab = (p) => {
  const b = readFileSync(join(V, p));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
};
const gc = () => {
  globalThis.gc?.();
  globalThis.gc?.();
};
const heap = () => {
  gc();
  return process.memoryUsage().heapUsed;
};
const MB = (x) => Math.round((x / 1048576) * 10) / 10;
async function olc(fn, n = 3) {
  const t = [];
  let r;
  for (let i = 0; i < n; i++) {
    const s = performance.now();
    r = await fn();
    t.push(performance.now() - s);
  }
  t.sort((a, b) => a - b);
  return { ms: Math.round(t[Math.floor(n / 2)]), min: Math.round(t[0]), r };
}
const out = {};

// ---- Belge (~300 sayfa)
{
  const data = ab("buyuk-belge.docx");
  const h0 = heap();
  const o = await olc(() => B.openDocument("buyuk-belge.docx", data));
  const doc = o.r.doc;
  const h1 = heap();
  const w = await olc(() => B.writeDocx(doc));
  const wo = await olc(() => B.writeOdt(doc), 1);
  const json = await olc(() => JSON.stringify({ content: doc.content, settings: doc.settings }), 5);
  const val = await olc(() => B.validateContent(doc.content), 3);
  out.belge = { dosyaKB: Math.round(data.byteLength / 1024), açMs: o.ms, docxYazMs: w.ms, odtYazMs: wo.ms, şemaDoğrulaMs: val.ms, taslakJSON_KB: Math.round(json.r.length / 1024), taslakStringifyMs: json.ms, modelBellekMB: MB(h1 - h0) };
  console.log("belge", out.belge);
}
// ---- Tablo 200k satır
{
  const data = ab("buyuk-tablo.xlsx");
  const h0 = heap();
  const o = await olc(() => T.openBook("buyuk-tablo.xlsx", data), 1);
  const wb = o.r.wb;
  const h1 = heap();
  const w = await olc(() => T.writeXlsx(wb), 1);
  const wc = await olc(() => T.writeBook(wb, wb.sheets[0], "csv"), 1);
  const wods = await olc(() => T.writeBook(wb, wb.sheets[0], "ods"), 1);
  const rc = await olc(() => wb.engine.recalc(true), 1);
  out.tablo200k = { dosyaMB: MB(data.byteLength), açMs: o.ms, xlsxYazMs: w.ms, xlsxÇıktıMB: MB(w.r.length), csvYazMs: wc.ms, odsYazMs: wods.ms, tamHesapMs: rc.ms, modelBellekMB: MB(h1 - h0) };
  console.log("tablo200k", out.tablo200k);
}
// ---- Formül zinciri 10k (+ bellekte 100k zincir: özyineleme taşması var mı?)
{
  const data = ab("zincir-10k.xlsx");
  const o = await olc(() => T.openBook("zincir-10k.xlsx", data), 3);
  const wb = o.r.wb;
  const s = wb.sheets[0];
  const tam = await olc(() => wb.engine.recalc(true), 5);
  let k = 2;
  const artimli = await olc(() => {
    wb.setInputRaw(s, 0, 0, k++);
    wb.engine.recalc();
  }, 5);
  const dogru = wb.getValue(s, 9999, 0) === k - 1 + 9999;
  const wb2 = new T.WorkbookModel(true);
  const s2 = wb2.addSheet("Z");
  wb2.engine.autoCalc = false;
  wb2.setInputRaw(s2, 0, 0, 1);
  for (let r = 1; r < 100000; r++) wb2.setInputRaw(s2, r, 0, { formula: `=A${r}+1` });
  wb2.engine.autoCalc = true;
  let derin;
  try {
    const t = performance.now();
    wb2.engine.recalc(true);
    const t2 = performance.now();
    wb2.setInputRaw(s2, 0, 0, 5);
    wb2.engine.recalc();
    derin = { tamHesapMs: Math.round(t2 - t), artımlıMs: Math.round(performance.now() - t2), sonDeğer: wb2.getValue(s2, 99999, 0) };
  } catch (e) {
    derin = { hata: `${e.constructor.name}: ${e.message}` };
  }
  out.zincir10k = { açMs: o.ms, tamHesapMs: tam.ms, A1DeğişinceArtımlıMs: artimli.ms, doğru: dogru, zincir100k: derin };
  console.log("zincir", out.zincir10k);
}
// ---- Sunum 200 slayt
{
  const data = new Uint8Array(ab("buyuk-sunu.pptx"));
  const h0 = heap();
  const o = await olc(() => S.readPresentationFile(data, "buyuk-sunu.pptx"), 3);
  const h1 = heap();
  const w = await olc(() => S.writePresentationFormat(o.r.pres, "pptx"), 3);
  const wo = await olc(() => S.writePresentationFormat(o.r.pres, "odp"), 1);
  out.sunum200 = { dosyaKB: Math.round(data.length / 1024), açMs: o.ms, pptxYazMs: w.ms, odpYazMs: wo.ms, slayt: o.r.pres.slides.length, modelBellekMB: MB(h1 - h0) };
  console.log("sunum200", out.sunum200);
}
mkdirSync(RAPOR, { recursive: true });
writeFileSync(join(RAPOR, "perf_node.json"), JSON.stringify(out, null, 1));
process.exit(0);
