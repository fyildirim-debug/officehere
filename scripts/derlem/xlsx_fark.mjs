// .xlsx yazıcı fark testi (P5.1): eski (exceljs) ve yeni (kendi) yazıcının çıktıları eşdeğer mi?
//
// Derlemdeki her Tablo dosyası (xlsx/xlsm/xls/ods/csv…) ve fast-check ile üretilen rastgele kitaplar
// iki yazıcıyla .xlsx olarak yazılır; iki çıktı kendi okuyucumuzla açılır ve modeller karşılaştırılır
// (hücre değer/formül, stiller, satır/sütun, birleştirme, not, köprü, doğrulama, koşullu biçim, tablo,
// grafik, resim, sayfa yapısı, adlar…). Fark yeni çıktıda özgün modele (yazmadan önceki) eşitse
// "iyileşme" sayılır (eski yazıcının kaybettiği bilgiyi yeni yazıcı koruyor); diğerleri "fark"tır.
// --lo ile LibreOffice (UNO) iki çıktıyı da açar, formülleri hesaplar ve sonuçlar karşılaştırılır.
//
// Kullanım: node scripts/derlem/xlsx_fark.mjs [--paketleme-yok] [--dosya <alt dize>] [--rastgele N]
//           [--tohum S] [--lo] [--cikti <dizin>] [--ek <dizin,dizin…>] (ek: derlem dışı dosyalar)
// Çıktı: .derlem/rapor/xlsx_fark.json ve xlsx_fark.md; fark varsa çıkış kodu 1.
import "./dom.mjs";
import { spawn, spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { createServer } from "node:net";
import { basename, join } from "node:path";
import { pathToFileURL } from "node:url";
import fc from "fast-check";
import { argumanlar, CIKTI, derlemDosyalari, GORSEL, KOK, PAKET, RAPOR, sure } from "./ortak.mjs";
import { paketle } from "./paketle.mjs";

const A = argumanlar();
const t0 = performance.now();
if (!A["paketleme-yok"]) await paketle({ sessiz: true });
const T = await import(pathToFileURL(join(PAKET, "tablo-io.mjs")).href);
const CIK = A.cikti ? String(A.cikti) : join(CIKTI, "xlsx_fark");
rmSync(CIK, { recursive: true, force: true });
mkdirSync(CIK, { recursive: true });

const toAB = (u8) => u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);

/** Seçili yazıcıyla yazar (yeni yazıcı katı kipte: hata olursa eskiye düşmez). */
async function yaz(wb, yazici) {
  const once = process.env.OH_XLSX_YAZICI;
  process.env.OH_XLSX_YAZICI = yazici;
  try {
    return await T.writeXlsx(wb);
  } finally {
    if (once === undefined) delete process.env.OH_XLSX_YAZICI;
    else process.env.OH_XLSX_YAZICI = once;
  }
}

/* ------------------------------------------------------------ model izi */

const sirali = (o) => {
  if (Array.isArray(o)) return o.map(sirali);
  if (o && typeof o === "object") return Object.fromEntries(Object.keys(o).sort().filter((k) => o[k] !== undefined).map((k) => [k, sirali(o[k])]));
  return o;
};
const idsiz = (o) => {
  if (Array.isArray(o)) return o.map(idsiz);
  if (o && typeof o === "object") return Object.fromEntries(Object.entries(o).filter(([k]) => k !== "id").map(([k, v]) => [k, idsiz(v)]));
  return o;
};
function deger(v) {
  if (v === null || v === undefined) return null;
  if (T.isError(v)) return { hata: v.code };
  if (typeof v === "number") return Number.isFinite(v) ? Math.round(v * 1e9) / 1e9 : String(v);
  return v;
}
const anahtar = (k) => `${T.keyRow(k)},${T.keyCol(k)}`;
const ozet = (s) => createHash("sha1").update(String(s)).digest("hex").slice(0, 12);

/** Modelin karşılaştırılabilir izi: düz, sıralı, kimliksiz nesne. */
function iz(wb) {
  wb.engine.recalc();
  const stil = (id) => JSON.stringify(sirali(wb.styles.get(id)));
  const sayfaAdi = (id) => wb.sheets.find((s) => s.id === id)?.name;
  return sirali({
    aktif: wb.activeIndex,
    ozellik: { baslik: wb.props.title || undefined, konu: wb.props.subject || undefined, yazar: wb.props.author || undefined, sirket: wb.props.company || undefined },
    tema: wb.theme,
    adlar: wb.engine
      .getNames()
      .map((n) => ({ ad: n.name, formul: n.formula, sayfa: n.sheet !== undefined ? sayfaAdi(n.sheet) : undefined }))
      .sort((a, b) => `${a.ad}|${a.sayfa}`.localeCompare(`${b.ad}|${b.sayfa}`)),
    koruma: wb.protection ?? undefined,
    sayfalar: wb.sheets.map((s) => {
      const hucre = {};
      wb.engine.forEachCell(s.id, (r, c) => {
        const f = wb.getFormula(s, r, c);
        hucre[`${r},${c}`] = f ? { f, v: deger(wb.getValue(s, r, c)) } : { v: deger(wb.getValue(s, r, c)) };
      });
      const harita = (m, fn = (v) => v) => Object.fromEntries([...m].map(([k, v]) => [anahtar(k), fn(v)]));
      const satirSutun = (m) => Object.fromEntries([...m].map(([k, v]) => [k, stil(v)]));
      const eksen = (a) => ({ varsayilan: a.defaultSize, ozel: a.customItems(), elle: a.manualItems?.(), gizli: a.hiddenItems(), duzey: a.levelItems(), daraltilmis: a.collapsedItems() });
      return {
        ad: s.name,
        gizli: s.hidden || undefined,
        sekme: s.tabColor,
        hucre,
        stil: harita(s.styles, stil),
        satirStil: satirSutun(s.rowStyles),
        sutunStil: satirSutun(s.colStyles),
        satir: eksen(s.rows),
        sutun: eksen(s.cols),
        birlesik: [...s.merges].map((m) => `${m.r1},${m.c1},${m.r2},${m.c2}`).sort(),
        not: harita(s.notes, (n) => n.text),
        kopru: harita(s.links),
        dogrulama: idsiz(s.validations),
        kosullu: idsiz(s.condFormats),
        filtre: s.autoFilter ?? undefined,
        tablo: idsiz(s.tables),
        grafik: idsiz(s.charts),
        resim: s.images.map((i) => ({ ...idsiz(i), src: ozet(i.src) })),
        sekil: idsiz(s.shapes),
        ozetTablo: idsiz(s.pivots),
        mini: idsiz(s.sparklines),
        senaryo: idsiz(s.scenarios),
        veriTablosu: idsiz(s.dataTables),
        dilimleyici: idsiz(s.slicers),
        yoksay: [...s.ignoredErrors].sort((a, b) => a - b),
        donmus: s.frozen,
        kilavuz: s.showGridlines,
        baslik: s.showHeadings,
        yakinlik: s.zoom,
        gorunum: s.viewMode,
        sayfa: s.page,
        secim: s.selection.active,
        korumali: s.protected || undefined,
      };
    }),
  });
}

/** İki izin farklı yolları (yaprak düzeyinde, ilk 200). */
function farklar(a, b, yol = "", out = []) {
  if (out.length >= 200) return out;
  if (a === b) return out;
  const nesne = (x) => x && typeof x === "object";
  if (!nesne(a) || !nesne(b) || Array.isArray(a) !== Array.isArray(b)) {
    if (JSON.stringify(a) !== JSON.stringify(b)) out.push(yol);
    return out;
  }
  if (Array.isArray(a) && a.length !== b.length) {
    out.push(yol);
    return out;
  }
  for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) farklar(a[k], b[k], yol ? `${yol}.${k}` : k, out);
  return out;
}
function al(o, yol) {
  let x = o;
  for (const k of yol.split(".")) x = x?.[k];
  return x;
}

/**
 * Eski ve yeni çıktının izlerini karşılaştırır. Fark yeni çıktıda özgün modele eşitse iyileşmedir;
 * yükleme uyarıları ve sayfa dışı ayrıntılar ayrıca raporlanır.
 */
function karsilastir(ozgun, eski, yeni) {
  const out = { fark: [], iyilesme: [] };
  for (const yol of farklar(eski, yeni)) {
    const e = al(eski, yol);
    const y = al(yeni, yol);
    const o = al(ozgun, yol);
    if (JSON.stringify(y) === JSON.stringify(o)) {
      out.iyilesme.push(yol);
      continue;
    }
    // Liste: yeni, eskinin tüm öğelerini ve özgünde karşılığı olan fazlasını içeriyorsa (eski yazıcının
    // düşürdüğü öğeler, ör. kopya/benzersiz değer kuralları) iyileşmedir
    if (Array.isArray(e) && Array.isArray(y) && Array.isArray(o) && y.length > e.length && y.length <= o.length) {
      const ys = y.map((x) => JSON.stringify(x));
      if (e.every((x) => ys.includes(JSON.stringify(x)))) {
        out.iyilesme.push(`${yol} (+${y.length - e.length} öğe)`);
        continue;
      }
    }
    out.fark.push({ yol, eski: JSON.stringify(e)?.slice(0, 160), yeni: JSON.stringify(y)?.slice(0, 160), ozgun: JSON.stringify(o)?.slice(0, 160) });
  }
  return out;
}

async function cift(ad, wb, sonuc) {
  const ozgun = iz(wb);
  const t1 = performance.now();
  const eskiB = await yaz(wb, "eski");
  const t2 = performance.now();
  const yeniB = await yaz(wb, "yeni");
  const t3 = performance.now();
  const dosya = ad.replace(/[^\p{L}\p{N}._-]+/gu, "_");
  const eskiYol = join(CIK, `${dosya}.eski.xlsx`);
  const yeniYol = join(CIK, `${dosya}.yeni.xlsx`);
  writeFileSync(eskiYol, eskiB);
  writeFileSync(yeniYol, yeniB);
  const yeni = iz((await T.openBook("yeni.xlsx", toAB(yeniB))).wb);
  let eski;
  let eskiAcilamadi;
  try {
    eski = iz((await T.openBook("eski.xlsx", toAB(eskiB))).wb);
  } catch (e) {
    eskiAcilamadi = `${e?.message ?? e} ${e?.detail ?? ""}`.trim();
  }
  // Eski çıktı açılamıyorsa (eski yazıcının hatası) yeni çıktı yalnız özgünle karşılaştırılır (bilgi amaçlı)
  const k = eski ? karsilastir(ozgun, eski, yeni) : { fark: [], iyilesme: ["(eski çıktı açılamadı)"], ozgunFarki: farklar(ozgun, yeni).slice(0, 30), eskiAcilamadi };
  sonuc.push({ ad, eskiMs: Math.round(t2 - t1), yeniMs: Math.round(t3 - t2), eskiKB: Math.round(eskiB.length / 1024), yeniKB: Math.round(yeniB.length / 1024), eskiYol, yeniYol, ...k });
  return k;
}

/* ------------------------------------------------------------ rastgele kitaplar */

const PNG = readFileSync(join(GORSEL, "logo.png")).toString("base64");
const ck = (r, c) => r * 16384 + c;

const metinA = fc.oneof(
  fc.constantFrom("", " ", "a", "Ankara", "İSTANBUL", "ığüşöç ĞÜŞİÖÇ", "<&>\"'", "x\ny", " baştaki", "sondaki ", "=A1", "_x0041_", "#YOK!", "1e5", "007", "\t", "a\u0001b", "😀"),
  fc.string({ maxLength: 12 }),
);
const sayiA = fc.oneof(fc.integer({ min: -1e6, max: 1e6 }), fc.double({ min: -1e9, max: 1e9, noNaN: true, noDefaultInfinity: true }), fc.constantFrom(0, -0, 0.1, 1e-7, 123456789012, 45292.5));
const formulA = fc.constantFrom("=A1+1", "=SUM(A1:B3)", "=IF(A1>0,\"evet\",\"hayır\")", "=1/0", "=CONCAT(\"a\";\"b\")", "=SEQUENCE(3)", "=SORT(A1:A4)", "=XLOOKUP(1,A1:A5,B1:B5)", "=TRUE()", "=NA()", "=LET(x,2,x*3)", "=UNIQUE(A1:A6)", "=TODAY()-TODAY()", "=B1&\"x\"");
const girdiA = fc.oneof({ weight: 3, arbitrary: sayiA }, { weight: 3, arbitrary: metinA }, { weight: 1, arbitrary: fc.boolean() }, { weight: 2, arbitrary: formulA.map((f) => ({ formula: f })) });
const renkA = fc.constantFrom("#FF0000", "#00B050", "#0070C0", "#FFFF00", "#000000", "#FFFFFF", "#7F7F7F");
const kenarA = fc.option(fc.record({ style: fc.constantFrom("thin", "medium", "thick", "dashed", "dotted", "double", "hair"), color: renkA }), { nil: undefined });
const stilA = fc.record(
  {
    font: fc.constantFrom("Arial", "Times New Roman", "Calibri Light"),
    size: fc.constantFrom(8, 10, 10.5, 14, 20),
    bold: fc.boolean(),
    italic: fc.boolean(),
    underline: fc.constantFrom("single", "double"),
    strike: fc.boolean(),
    color: renkA,
    fill: renkA,
    pattern: fc.record({ type: fc.constantFrom("gray125", "darkGrid", "lightDown"), color: renkA }),
    hAlign: fc.constantFrom("left", "center", "right", "justify", "fill"),
    vAlign: fc.constantFrom("top", "middle", "bottom"),
    wrap: fc.boolean(),
    shrink: fc.boolean(),
    indent: fc.integer({ min: 0, max: 4 }),
    rotation: fc.constantFrom(0, 45, -30, 90, 255),
    numFmt: fc.constantFrom("0.00", "0%", "#,##0", "dd.mm.yyyy", '#,##0.00 "₺"', "@", "0.00E+00", "[Red]0;-0", "General"),
    bTop: kenarA,
    bBottom: kenarA,
    bLeft: kenarA,
    bRight: kenarA,
    locked: fc.constant(false),
    hidden: fc.boolean(),
    checkbox: fc.boolean(),
  },
  { requiredKeys: [] },
);
const kuralA = fc.oneof(
  fc.record({ type: fc.constant("cellIs"), operator: fc.constantFrom("greaterThan", "between", "equal"), values: fc.constantFrom(["5"], ["1", "10"], ["=$A$1"]) }),
  fc.record({ type: fc.constantFrom("containsText", "notContainsText", "beginsWith", "endsWith"), text: fc.constantFrom("a", "ab", "x y") }),
  fc.record({ type: fc.constantFrom("duplicateValues", "uniqueValues", "containsBlanks", "notContainsBlanks", "containsErrors") }),
  fc.record({ type: fc.constant("top10"), rank: fc.integer({ min: 1, max: 20 }), percent: fc.boolean(), bottom: fc.boolean() }),
  fc.record({ type: fc.constant("aboveAverage"), below: fc.boolean() }),
  fc.record({ type: fc.constant("expression"), values: fc.constantFrom(["=A1>2"], ["=MOD(ROW(),2)=0"]) }),
  fc.record({ type: fc.constant("timePeriod"), timePeriod: fc.constantFrom("today", "yesterday", "last7Days", "thisMonth", "nextWeek") }),
  fc.record({ type: fc.constant("dataBar"), color: renkA }),
  fc.record({ type: fc.constant("colorScale"), colors: fc.constantFrom(["#F8696B", "#63BE7B"], ["#F8696B", "#FFEB84", "#63BE7B"]) }),
  fc.record({ type: fc.constant("iconSet"), iconSet: fc.constantFrom("3Arrows", "3TrafficLights1", "3Triangles", "5Boxes", "4Rating", "3Stars"), reverseIcons: fc.boolean(), iconOnly: fc.boolean() }),
);
const aralikA = fc.record({ r1: fc.integer({ min: 0, max: 20 }), c1: fc.integer({ min: 0, max: 6 }), h: fc.integer({ min: 0, max: 5 }), w: fc.integer({ min: 0, max: 3 }) }).map(({ r1, c1, h, w }) => ({ r1, c1, r2: r1 + h, c2: c1 + w }));
const sayfaA = fc.record({
  ad: fc.constantFrom("Sayfa", "Veri Ç", "Özet&<>", "Liste", "Sayfa", "Liste", "Veri Ç", "O'Brien"),
  hucreler: fc.array(fc.record({ r: fc.integer({ min: 0, max: 25 }), c: fc.integer({ min: 0, max: 7 }), v: girdiA }), { maxLength: 40 }),
  uzak: fc.option(fc.record({ r: fc.integer({ min: 1000, max: 1048575 }), c: fc.integer({ min: 0, max: 16383 }), v: girdiA }), { nil: undefined }),
  stiller: fc.array(stilA, { maxLength: 6 }),
  kendiStil: fc.array(fc.record({ r: fc.integer({ min: 0, max: 27 }), c: fc.integer({ min: 0, max: 8 }), s: fc.nat(6) }), { maxLength: 20 }),
  satirStil: fc.array(fc.record({ i: fc.integer({ min: 0, max: 30 }), s: fc.nat(6) }), { maxLength: 4 }),
  sutunStil: fc.array(fc.record({ i: fc.integer({ min: 0, max: 9 }), s: fc.nat(6) }), { maxLength: 4 }),
  satirBoyut: fc.array(fc.record({ i: fc.integer({ min: 0, max: 40 }), px: fc.integer({ min: 0, max: 80 }), elle: fc.boolean(), gizli: fc.boolean(), duzey: fc.integer({ min: 0, max: 3 }), daralt: fc.boolean() }), { maxLength: 6 }),
  sutunBoyut: fc.array(fc.record({ i: fc.integer({ min: 0, max: 12 }), px: fc.integer({ min: 0, max: 300 }), gizli: fc.boolean(), duzey: fc.integer({ min: 0, max: 2 }) }), { maxLength: 6 }),
  birlesik: fc.array(fc.record({ r: fc.integer({ min: 31, max: 37 }), c: fc.integer({ min: 0, max: 6 }), h: fc.integer({ min: 0, max: 2 }), w: fc.integer({ min: 0, max: 2 }), deger: fc.option(girdiA, { nil: undefined }), stil: fc.option(fc.nat(6), { nil: undefined }) }), { maxLength: 3 }),
  notlar: fc.array(fc.record({ r: fc.integer({ min: 0, max: 40 }), c: fc.integer({ min: 0, max: 9 }), t: metinA }), { maxLength: 4 }),
  kopruler: fc.array(fc.record({ r: fc.integer({ min: 0, max: 40 }), c: fc.integer({ min: 0, max: 9 }), l: fc.constantFrom("https://ornek.com/a?b=1&c=2", "#Liste!A1", "mailto:a@b.c", "#'O''Brien'!B2") }), { maxLength: 4 }),
  dogrulama: fc.array(fc.record({ rule: fc.oneof(fc.record({ type: fc.constant("list"), formula1: fc.constantFrom("a;b;c", "=$A$1:$A$3"), showDropdown: fc.boolean() }), fc.record({ type: fc.constantFrom("whole", "decimal"), operator: fc.constantFrom("between", "greaterThan"), formula1: fc.constantFrom("1", "=A1"), formula2: fc.constant("10"), error: fc.constantFrom("Hata", "Yanlış <değer>") })), ranges: fc.array(aralikA, { minLength: 1, maxLength: 2 }) }), { maxLength: 2 }),
  kosullu: fc.array(fc.record({ kural: kuralA, ranges: fc.array(aralikA, { minLength: 1, maxLength: 2 }), stil: fc.option(stilA, { nil: undefined }) }), { maxLength: 4 }),
  filtre: fc.boolean(),
  tablo: fc.option(fc.record({ c: fc.integer({ min: 0, max: 3 }), w: fc.integer({ min: 1, max: 3 }), h: fc.integer({ min: 0, max: 4 }), baslik: fc.boolean(), toplam: fc.boolean(), stil: fc.constantFrom("TableStyleMedium2", "TableStyleLight1", "") }), { nil: undefined }),
  resim: fc.option(fc.record({ r: fc.integer({ min: 0, max: 10 }), c: fc.integer({ min: 0, max: 5 }), dx: fc.integer({ min: 0, max: 30 }), dy: fc.integer({ min: 0, max: 10 }), w: fc.integer({ min: 5, max: 300 }), h: fc.integer({ min: 5, max: 200 }) }), { nil: undefined }),
  donmus: fc.record({ rows: fc.integer({ min: 0, max: 3 }), cols: fc.integer({ min: 0, max: 2 }) }),
  yakinlik: fc.constantFrom(100, 75, 150),
  kilavuz: fc.boolean(),
  basliklar: fc.boolean(),
  gorunum: fc.constantFrom("normal", "pageLayout", "pageBreak"),
  sayfa: fc.record({ orientation: fc.constantFrom("portrait", "landscape"), paper: fc.constantFrom("A4", "A3", "Letter", "A5", "Legal"), printArea: fc.option(aralikA, { nil: undefined }), titleRows: fc.option(fc.constant({ r1: 0, r2: 1 }), { nil: undefined }), titleCols: fc.option(fc.constant({ c1: 0, c2: 0 }), { nil: undefined }), fitMode: fc.constantFrom(undefined, "sheet", "cols", "rows"), header: fc.constantFrom(undefined, "Başlık &[Sayfa]", "A & B"), footer: fc.constantFrom(undefined, "&[Sayfa] / &[Sayfalar]"), printGridlines: fc.boolean(), printHeadings: fc.boolean(), blackAndWhite: fc.boolean(), draft: fc.boolean(), pageOrder: fc.constantFrom(undefined, "overThenDown") }),
  korumali: fc.boolean(),
  gizli: fc.boolean(),
  sekme: fc.option(renkA, { nil: undefined }),
  veriTablosu: fc.boolean(),
});
const kitapA = fc.record({ sayfalar: fc.array(sayfaA, { minLength: 1, maxLength: 3 }), aktif: fc.nat(2), adlar: fc.array(fc.record({ ad: fc.constantFrom("Oran", "Liste_1", "Toplam"), f: fc.constantFrom("=0,18", "=$A$1:$A$3", "=SUM($B$1:$B$2)"), yerel: fc.boolean() }), { maxLength: 2 }), baslik: fc.option(fc.constantFrom("Rapor", "Ç<ö>"), { nil: undefined }) });

/** Tanımdaki kesme işaretli sayfa adı (ve ona köprü) tipografik kesmeyle: eski yazıcının açılabilir çıktısı için. */
const kesmesiz = (x) =>
  typeof x === "string" ? x.replace(/O'+Brien/g, "O’Brien") : Array.isArray(x) ? x.map(kesmesiz) : x && typeof x === "object" ? Object.fromEntries(Object.entries(x).map(([k, v]) => [k, kesmesiz(v)])) : x;

/** Rastgele kitap tanımından model kurar (geçerli, uygulamanın da üretebileceği biçimde). */
function kitapKur(spec) {
  const wb = new T.WorkbookModel(true);
  const adlar = new Set();
  spec.sayfalar.forEach((sp, si) => {
    let ad = sp.ad;
    while (adlar.has(ad.toLocaleUpperCase("tr-TR"))) ad += String(si);
    adlar.add(ad.toLocaleUpperCase("tr-TR"));
    const s = wb.addSheet(ad);
    const stiller = sp.stiller.map((st) => wb.styles.intern(st));
    const st = (i) => (stiller.length ? stiller[i % stiller.length] : 0);
    for (const h of sp.hucreler) wb.setInputRaw(s, h.r, h.c, h.v);
    if (sp.uzak) wb.setInputRaw(s, sp.uzak.r, sp.uzak.c, sp.uzak.v);
    for (const h of sp.kendiStil) s.styles.set(ck(h.r, h.c), st(h.s));
    for (const x of sp.satirStil) s.rowStyles.set(x.i, st(x.s));
    for (const x of sp.sutunStil) s.colStyles.set(x.i, st(x.s));
    for (const x of sp.satirBoyut) {
      if (x.px !== 20 || x.elle) s.rows.set(x.i, x.px, x.elle);
      if (x.gizli) s.rows.setHidden(x.i, true);
      if (x.duzey) s.rows.setLevel(x.i, x.duzey);
      if (x.daralt) s.rows.setCollapsed(x.i, true);
    }
    for (const x of sp.sutunBoyut) {
      if (x.px !== 64) s.cols.set(x.i, x.px, true);
      if (x.gizli) s.cols.setHidden(x.i, true);
      if (x.duzey) s.cols.setLevel(x.i, x.duzey);
    }
    for (const m of sp.birlesik) {
      if (!m.h && !m.w) continue;
      const r = { r1: m.r, c1: m.c, r2: m.r + m.h, c2: m.c + m.w };
      if (s.merges.some((x) => x.r1 <= r.r2 && r.r1 <= x.r2 && x.c1 <= r.c2 && r.c1 <= x.c2)) continue;
      s.merges.push(r);
      if (m.deger !== undefined) wb.setInputRaw(s, m.r, m.c, m.deger);
      if (m.stil !== undefined) s.styles.set(ck(m.r, m.c), st(m.stil));
    }
    for (const n of sp.notlar) s.notes.set(ck(n.r, n.c), { text: n.t || "not" });
    for (const k of sp.kopruler) s.links.set(ck(k.r, k.c), k.l);
    for (const d of sp.dogrulama) s.validations.push({ id: wb.newId("dv"), rule: d.rule, ranges: d.ranges });
    for (const k of sp.kosullu) s.condFormats.push({ id: wb.newId("cf"), ...k.kural, ranges: k.ranges, ...(k.stil && !["dataBar", "colorScale", "iconSet"].includes(k.kural.type) ? { style: k.stil } : {}) });
    if (sp.filtre) s.autoFilter = { range: { r1: 0, c1: 0, r2: 20, c2: 4 }, columns: {} };
    if (sp.tablo) {
      const t = sp.tablo;
      const r1 = 40;
      const r2 = r1 + t.h + (t.baslik ? 1 : 0) + (t.toplam ? 1 : 0);
      const columns = Array.from({ length: t.w }, (_, i) => `Sütun ${i + 1}`);
      if (t.baslik) columns.forEach((c, i) => wb.setInputRaw(s, r1, t.c + i, c));
      for (let r = r1 + (t.baslik ? 1 : 0); r < r2 - (t.toplam ? 1 : 0); r++) for (let i = 0; i < t.w; i++) wb.setInputRaw(s, r, t.c + i, r * 10 + i);
      if (r2 > r1) s.tables.push({ id: wb.newId("t"), name: `Tablo${si + 1}`, ref: { r1, c1: t.c, r2, c2: t.c + t.w - 1 }, style: t.stil, headerRow: t.baslik, totalRow: t.toplam, bandedRows: true, bandedCols: false, firstColumn: false, lastColumn: false, filterButton: true, columns });
    }
    if (sp.resim) s.images.push({ id: wb.newId("img"), name: "Resim 1", src: `data:image/png;base64,${PNG}`, anchor: { row: sp.resim.r, col: sp.resim.c, dx: sp.resim.dx, dy: sp.resim.dy }, width: sp.resim.w, height: sp.resim.h });
    s.frozen = sp.donmus;
    s.zoom = sp.yakinlik;
    s.showGridlines = sp.kilavuz;
    s.showHeadings = sp.basliklar;
    s.viewMode = sp.gorunum;
    Object.assign(s.page, Object.fromEntries(Object.entries(sp.sayfa).filter(([, v]) => v !== undefined)));
    s.protected = sp.korumali;
    s.tabColor = sp.sekme;
    if (sp.veriTablosu) {
      wb.setInputRaw(s, 45, 0, { formula: "=A46*2" });
      for (let i = 0; i < 3; i++) wb.setInputRaw(s, 46 + i, 0, i + 1);
      s.dataTables.push({ id: wb.newId("dt"), ref: { r1: 46, c1: 1, r2: 48, c2: 1 }, colInput: { row: 45, col: 0 } });
    }
  });
  if (spec.sayfalar.length > 1 && spec.sayfalar[spec.sayfalar.length - 1].gizli) wb.sheets[wb.sheets.length - 1].hidden = true;
  wb.activeIndex = Math.min(spec.aktif, wb.sheets.length - 1);
  if (wb.sheets[wb.activeIndex].hidden) wb.activeIndex = 0;
  for (const n of spec.adlar) {
    try {
      wb.engine.defineName(n.ad, n.f, n.yerel ? wb.sheets[0].id : undefined);
    } catch {
      /* geçersiz ad */
    }
  }
  if (spec.baslik) wb.props.title = spec.baslik;
  wb.syncTables();
  wb.engine.recalc(true);
  return wb;
}

/* ------------------------------------------------------------ çalıştır */

const sonuc = [];
const hatalar = [];
const dosyalar = derlemDosyalari({ uygulama: "tablo", suzgec: A.dosya });
for (const d of String(A.ek ?? "").split(",").filter(Boolean)) {
  if (!existsSync(d)) continue;
  for (const f of readdirSync(d).sort()) if (/\.(xlsx|xlsm|xltx|xltm|xls|ods|csv)$/i.test(f) && (!A.dosya || f.includes(A.dosya))) dosyalar.push({ yol: join(d, f), anahtar: `ek/${f}` });
}
const atlanan = [];
for (const f of dosyalar) {
  let wb;
  try {
    wb = (await T.openBook(basename(f.yol), toAB(readFileSync(f.yol)))).wb;
  } catch (e) {
    // Açılamayan kaynak (şifreli, bozuk): yazıcıyla ilgisi yok
    atlanan.push({ ad: f.anahtar, neden: String(e?.message ?? e).slice(0, 200) });
    continue;
  }
  try {
    await cift(f.anahtar, wb, sonuc);
  } catch (e) {
    hatalar.push({ ad: f.anahtar, hata: String(e?.stack ?? e).slice(0, 600) });
  }
}
const N = Number(A.rastgele ?? 150);
const tohum = Number(A.tohum ?? 20261015);
let rastgeleFark = null;
if (N > 0) {
  let i = 0;
  try {
    await fc.assert(
      fc.asyncProperty(kitapA, async (spec) => {
        const ad = `rastgele-${String(i++).padStart(3, "0")}`;
        const yerel = [];
        const k = await cift(ad, kitapKur(spec), yerel);
        if (k.fark.length) {
          rastgeleFark = { spec, ...yerel[0] };
          return false;
        }
        // Eski çıktı açılamadıysa (kesme işaretli sayfa adı) aynı kitap kesmesiz adla da karşılaştırılır
        if (k.eskiAcilamadi) {
          const k2 = await cift(`${ad} (kesmesiz)`, kitapKur(kesmesiz(spec)), yerel);
          if (k2.fark.length) {
            rastgeleFark = { spec: kesmesiz(spec), ...yerel[1] };
            return false;
          }
        }
        sonuc.push(...yerel);
        return true;
      }),
      { numRuns: N, seed: tohum, endOnFailure: !A.kucult },
    );
  } catch (e) {
    const neden = e?.cause ? `\nNeden: ${String(e.cause?.stack ?? e.cause).slice(0, 2000)}` : "";
    hatalar.push({ ad: "rastgele (fast-check)", hata: `${String(e?.message ?? e).slice(0, 6000)}${neden}` });
    if (rastgeleFark) sonuc.push(rastgeleFark);
  }
}

/* ------------------------------------------------------------ LibreOffice (UNO) */

let lo = null;
if (A.lo) {
  const var_ = spawnSync("soffice", ["--version"], { stdio: "ignore" }).status === 0 && spawnSync("python3", ["-c", "import uno"], { stdio: "ignore" }).status === 0;
  if (!var_) lo = { atlandi: "soffice ya da python3-uno yok" };
  else {
    const port = await new Promise((res) => {
      const s = createServer();
      s.listen(0, "127.0.0.1", () => {
        const p = s.address().port;
        s.close(() => res(p));
      });
    });
    const profil = join(CIKTI, "lo-profil-fark");
    mkdirSync(profil, { recursive: true });
    const proc = spawn("soffice", ["--headless", "--invisible", "--nologo", "--norestore", "--nodefault", `--accept=socket,host=127.0.0.1,port=${port};urp;`, `-env:UserInstallation=${pathToFileURL(profil).href}`], { stdio: "ignore", detached: true });
    const dene = `import uno\nl=uno.getComponentContext()\nl.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",l).resolve("uno:socket,host=127.0.0.1,port=${port};urp;StarOffice.ComponentContext")`;
    let hazir = false;
    for (let i = 0; i < 90 && !hazir; i++) {
      hazir = spawnSync("python3", ["-c", dene], { stdio: "ignore" }).status === 0;
      if (!hazir) await new Promise((r) => setTimeout(r, 1000));
    }
    try {
      if (!hazir) throw new Error("LibreOffice UNO dinleyicisi hazır olmadı");
      const args = sonuc.flatMap((s) => [`${s.ad}#eski=${s.eskiYol}`, `${s.ad}#yeni=${s.yeniYol}`]);
      const r = spawnSync("python3", [join(KOK, "scripts/derlem/gercek_lo.py"), ...args], { cwd: KOK, env: { ...process.env, LO_PORT: String(port) }, maxBuffer: 1 << 30, encoding: "utf8" });
      if (r.status !== 0) throw new Error(`gercek_lo.py başarısız: ${r.stderr?.slice(0, 400)}`);
      const deg = JSON.parse(r.stdout);
      const esit = (a, b) => (typeof a === "number" && typeof b === "number" ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a)) : a === b);
      lo = { dosya: 0, acilamayan: [], farkli: [] };
      for (const s of sonuc) {
        const e = deg[`${s.ad}#eski`];
        const y = deg[`${s.ad}#yeni`];
        lo.dosya++;
        if (!y || y.__hata__) {
          lo.acilamayan.push({ ad: s.ad, hata: y?.__hata__ ?? "sonuç yok" });
          continue;
        }
        if (!e || e.__hata__) continue;
        const fark = [];
        for (const sh of new Set([...Object.keys(e), ...Object.keys(y)]))
          for (const c of new Set([...Object.keys(e[sh] ?? {}), ...Object.keys(y[sh] ?? {})])) {
            // Geçici işlevler (BUGÜN, SLS…) iki açılış arasında değişebilir: yalnız eşit olmayanlar
            if (!esit(e[sh]?.[c], y[sh]?.[c])) fark.push(`${sh}!${c}: ${JSON.stringify(e[sh]?.[c])} → ${JSON.stringify(y[sh]?.[c])}`);
          }
        if (fark.length) lo.farkli.push({ ad: s.ad, fark: fark.slice(0, 10) });
      }
    } catch (e) {
      lo = { hata: String(e?.message ?? e) };
    } finally {
      const kapat = `import uno\nl=uno.getComponentContext()\nc=l.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",l).resolve("uno:socket,host=127.0.0.1,port=${port};urp;StarOffice.ComponentContext")\ntry:\n  c.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",c).terminate()\nexcept Exception:\n  pass`;
      spawnSync("python3", ["-c", kapat], { stdio: "ignore", timeout: 20000 });
      try {
        process.kill(-proc.pid, "SIGKILL");
      } catch {
        /* kapandı */
      }
    }
  }
}

/* ------------------------------------------------------------ rapor */

const farkli = sonuc.filter((s) => s.fark.length);
const iyi = sonuc.filter((s) => s.iyilesme.length);
const top = (k) => sonuc.reduce((a, s) => a + s[k], 0);
const L = [
  "# .xlsx yazıcı fark testi (eski exceljs ↔ yeni)",
  "",
  `${sonuc.length} kitap (${dosyalar.length} derlem dosyası, ${N} rastgele) · ${sure(performance.now() - t0)}`,
  "",
  `- Model farkı olan kitap: **${farkli.length}**`,
  `- Yalnız iyileşme (yeni çıktı özgüne eşit): ${iyi.length} kitap, ${iyi.reduce((a, s) => a + s.iyilesme.length, 0)} yol`,
  `- Eski çıktısı açılamayan: ${sonuc.filter((s) => s.eskiAcilamadi).length} kitap (kesmesiz eşi karşılaştırılamayan: ${sonuc.filter((s) => s.eskiAcilamadi && s.ad.endsWith("(kesmesiz)")).length})`,
  `- Hata: ${hatalar.length}; açılamayan kaynak (atlandı): ${atlanan.length}`,
  `- Toplam yazma süresi: eski ${sure(top("eskiMs"))}, yeni ${sure(top("yeniMs"))}; toplam boyut eski ${top("eskiKB")} KB, yeni ${top("yeniKB")} KB`,
];
if (lo) L.push(`- LibreOffice: ${lo.atlandi ?? lo.hata ?? `${lo.dosya} kitap, açılamayan ${lo.acilamayan.length}, hesap farkı ${lo.farkli.length}`}`);
L.push("");
for (const s of farkli) {
  L.push(`## ${s.ad}`, "");
  for (const d of s.fark.slice(0, 20)) L.push(`- \`${d.yol}\`: eski ${d.eski} · yeni ${d.yeni} · özgün ${d.ozgun}`);
  L.push("");
}
if (iyi.length) {
  L.push("## İyileşmeler (yeni çıktı özgün modele eşit)", "");
  for (const s of iyi) L.push(`- ${s.ad}: ${s.iyilesme.slice(0, 8).join(", ")}${s.iyilesme.length > 8 ? " …" : ""}`);
  L.push("");
}
if (hatalar.length) {
  L.push("## Hatalar", "");
  for (const h of hatalar) L.push(`- ${h.ad}: ${h.hata.split("\n").slice(0, 6).join(" | ")}`);
}
if (lo?.acilamayan?.length || lo?.farkli?.length) {
  L.push("", "## LibreOffice", "");
  for (const a of lo.acilamayan) L.push(`- açılamadı: ${a.ad} (${a.hata})`);
  for (const f of lo.farkli) L.push(`- hesap farkı: ${f.ad}: ${f.fark.join("; ")}`);
}
mkdirSync(RAPOR, { recursive: true });
writeFileSync(join(RAPOR, "xlsx_fark.json"), JSON.stringify({ sonuc, hatalar, atlanan, lo, rastgeleFark }, null, 1));
writeFileSync(join(RAPOR, "xlsx_fark.md"), L.join("\n") + "\n");
console.log(L.slice(0, 10).join("\n"));
if (farkli.length || hatalar.length || lo?.acilamayan?.length || lo?.farkli?.length) {
  console.log(`Ayrıntı: ${join(RAPOR, "xlsx_fark.md")}`);
  process.exit(1);
}
process.exit(0);
