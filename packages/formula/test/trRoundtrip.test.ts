/**
 * Türkçe işlev adı gidiş-dönüşü (rakip araştırması Ö13). LibreOffice'in iki açık hatasının bizde
 * olmadığını kanıtlar: 135020 (BÜYÜKHARF "i"yi "I" yapıyor) ve 142293 (Türkçe yerel ayarda kaydederken
 * işlev adı "İ" ile bozuluyor).
 *
 * Kayıtlı her işlev için: Türkçe yazım → kanonik (İngilizce) formül → XLSX dosya formülü (Tablo'nun
 * yazıcısının önek dönüşümü + XML kaçışı + UTF-8) → geri okuma → Türkçe yazım aynı olmalı; dosyaya
 * Türkçe karakter sızmamalı. BÜYÜKHARF / KÜÇÜKHARF / YAZIM.DÜZENİ'nde İ/ı/i/I doğruluğu. fast-check
 * ile rastgele harf büyüklüğü, rastgele iç içe formüller ve rastgele Türkçe metinler (1000'er yineleme).
 */
import fc from "fast-check";
import { describe, expect, it } from "vitest";
// Dosya formülü dönüşümü Tablo'nun XLSX yazıcı/okuyucusunun kullandığı işlevlerin kendisidir.
import { addXlfnPrefixes, normalizeFileFormula } from "../../../apps/tablo/src/io/xlsxFormula";
import { EN_LOCALE, ERROR_TO_TR, FormulaEngine, FUNCTIONS, fromLocalFormula, parseCellRef, parseFormula, toLocalFormula, type FnDef, type Scalar } from "../src";

const DEFS = [...FUNCTIONS.values()].sort((a, b) => a.name.localeCompare(b.name));
const trUst = (s: string) => s.toLocaleUpperCase("tr-TR");
const trAlt = (s: string) => s.toLocaleLowerCase("tr-TR");

/** XLSX'te <f> öğesinin metni olarak yazılıp okunması: XML kaçışı ve UTF-8. */
function dosyayaYazOku(formul: string): string {
  const xml = `<f>${formul.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</f>`;
  const geri = new TextDecoder().decode(new TextEncoder().encode(xml));
  return geri.slice(3, -4).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
}

/** Tür tam döngü: Türkçe formül → kanonik → dosya → kanonik → Türkçe. */
function dongu(tr: string) {
  const kanonik = fromLocalFormula(tr, "tr");
  const dosya = addXlfnPrefixes(kanonik.slice(1));
  const okunan = "=" + normalizeFileFormula(dosyayaYazOku(dosya));
  return { kanonik, dosya, okunan, geri: toLocalFormula(okunan, "tr") };
}

/** Metin sabitleri dışındaki karakterler (dosya formülünde yalnız ASCII olmalı). */
const sabitlerDisi = (f: string) => f.replace(/"(?:[^"]|"")*"/g, '""');

// Bağımsız değişken havuzu: [Türkçe yazım, kanonik yazım]
const ARG: [string, string][] = [
  ["A1", "A1"],
  ["1,5", "1.5"],
  ['"a;b,c ""İı"""', '"a;b,c ""İı"""'],
  ["Sayfa2!B2:C3", "Sayfa2!B2:C3"],
  ["DOĞRU", "TRUE"],
  ["#YOK", "#N/A"],
];
function ornek(d: FnDef): { tr: string; en: string } {
  const n = d.max === 0 ? 0 : Math.min(Math.max(d.min, 1), d.max, ARG.length);
  const a = ARG.slice(0, n);
  return { tr: `=${d.tr}(${a.map((x) => x[0]).join(";")})`, en: `=${d.name}(${a.map((x) => x[1]).join(",")})` };
}

describe("Türkçe işlev adları: kayıt", () => {
  it("370'ten fazla işlev; adlar tekil, büyük harfli, kanonik adlar ASCII", () => {
    expect(DEFS.length).toBeGreaterThanOrEqual(370);
    const sorun: string[] = [];
    const trAdlar = new Map<string, string>();
    for (const d of DEFS) {
      if (!/^[A-Z][A-Z0-9.]*$/.test(d.name)) sorun.push(`${d.name}: kanonik ad ASCII büyük harf değil`);
      if (!d.tr || d.tr !== trUst(d.tr)) sorun.push(`${d.name}: Türkçe ad (${d.tr}) Türkçe büyük harf değil`);
      if (/[a-z]/.test(d.tr)) sorun.push(`${d.name}: Türkçe adda küçük harf (${d.tr})`);
      const k = trUst(d.tr);
      if (trAdlar.has(k)) sorun.push(`${d.name}: Türkçe ad ${d.tr} ${trAdlar.get(k)} ile aynı`);
      trAdlar.set(k, d.name);
    }
    // Bir işlevin Türkçe adı başka bir işlevin İngilizce adıyla aynı olmamalı (yazım iki anlamlı olur)
    for (const d of DEFS) {
      const baska = FUNCTIONS.get(d.tr);
      if (baska && baska.name !== d.name) sorun.push(`${d.name}: Türkçe ad ${d.tr} başka bir işlevin (${baska.name}) İngilizce adı`);
    }
    expect(sorun).toEqual([]);
  });
});

describe("Türkçe işlev adları: tr → en → XLSX → tr", () => {
  it("her işlev: Türkçe yazım kanoniğe, dosyaya ve geri aynen döner; dosyada Türkçe karakter yok", () => {
    const sorun: string[] = [];
    for (const d of DEFS) {
      const { tr, en } = ornek(d);
      const r = dongu(tr);
      if (r.kanonik !== en) sorun.push(`${d.name}: tr→en ${tr} → ${r.kanonik} (beklenen ${en})`);
      if (r.okunan !== r.kanonik) sorun.push(`${d.name}: dosya gidiş-dönüşü ${r.kanonik} → ${r.dosya} → ${r.okunan}`);
      if (r.geri !== tr) sorun.push(`${d.name}: en→tr ${r.okunan} → ${r.geri} (beklenen ${tr})`);
      if (!/^[\x20-\x7e]*$/.test(sabitlerDisi(r.dosya))) sorun.push(`${d.name}: dosya formülünde ASCII dışı karakter: ${r.dosya}`);
      if (!r.dosya.startsWith(d.xlfn ? "_xlfn." : d.name + "(")) sorun.push(`${d.name}: dosya adı/öneki yanlış: ${r.dosya}`);
      try {
        const kok = parseFormula(r.okunan);
        if (kok.type !== "func" || kok.name !== d.name) sorun.push(`${d.name}: dosyadan okunan formül başka işleve çözüldü (${kok.type === "func" ? kok.name : kok.type})`);
      } catch (e) {
        sorun.push(`${d.name}: dosyadan okunan formül ayrıştırılamadı: ${(e as Error).message}`);
      }
    }
    expect(sorun).toEqual([]);
  });

  it("her işlev: küçük harfle Türkçe ad ve Türkçe yerel ayarda İngilizce ad aynı işleve çözülür", () => {
    const sorun: string[] = [];
    for (const d of DEFS) {
      const { tr, en } = ornek(d);
      const govde = tr.slice(1 + d.tr.length);
      for (const ad of [trAlt(d.tr), d.name.toLowerCase(), d.name, d.name.toLowerCase().replace(/i/g, "ı")]) {
        // "ı" yalnız İngilizce adın Türkçe klavyede yazılmış hâli: ad "I" içermiyorsa denenmez
        if (ad.includes("ı") && !d.name.includes("I")) continue;
        const k = fromLocalFormula(`=${ad}${govde}`, "tr");
        if (k !== en) sorun.push(`${d.name}: "${ad}" → ${k} (beklenen ${en})`);
      }
    }
    expect(sorun).toEqual([]);
  });

  it("bilinmeyen ad (kullanıcı tanımlı işlev) iki yönde de olduğu gibi kalır; dosyaya Türkçe büyük harfle yazılmaz", () => {
    for (const tr of ["=myudfi(A1)", "=Hesapla_i(1;2)", "=ışık(A1)", "=İŞLE(1)"]) {
      const r = dongu(tr);
      expect(r.geri, tr).toBe(tr);
      expect(r.okunan, tr).toBe(r.kanonik);
    }
    expect(addXlfnPrefixes(fromLocalFormula("=myudfi(A1)").slice(1))).toBe("myudfi(A1)");
  });

  it("Türkçe adlı LAMBDA: ekranda gösterilip yeniden girilen formül aynı ada çözülür", () => {
    const e = new FormulaEngine();
    const s = e.addSheet("Sayfa1");
    e.defineName("işle", fromLocalFormula("=LAMBDA(x;x*2)"));
    e.defineName("ıslak", fromLocalFormula("=LAMBDA(x;x+1)"));
    e.setFormula(s, 0, 0, fromLocalFormula("=işle(2)+ıslak(1)"));
    // Dosyadan okunan formül (kanonik, küçük harfle yazılmış ad)
    e.setFormula(s, 1, 0, "=işle(2)+ıslak(1)");
    e.recalc();
    expect(e.getValue(s, 0, 0)).toBe(6);
    expect(e.getValue(s, 1, 0)).toBe(6);
    // Hücre düzenlenip değiştirilmeden onaylanır: Türkçe gösterim → kanonik
    for (const r of [0, 1]) {
      const yeniden = fromLocalFormula(toLocalFormula(e.getFormula(s, r, 0)!, "tr"), "tr");
      e.setFormula(s, r, 1, yeniden);
      e.recalc();
      expect(e.getValue(s, r, 1), yeniden).toBe(6);
    }
  });

  it("İ/ı içeren adlar ve hata değerleri", () => {
    expect(fromLocalFormula("=indis(A1:B2;1;1)")).toBe("=INDEX(A1:B2,1,1)");
    expect(fromLocalFormula("=İNDİS(A1:B2;1;1)")).toBe("=INDEX(A1:B2,1,1)");
    expect(fromLocalFormula("=kırp(A1)")).toBe("=TRIM(A1)");
    expect(fromLocalFormula("=yazım.düzeni(A1)")).toBe("=PROPER(A1)");
    expect(fromLocalFormula("=eğerhata(1/0;#değer!)")).toBe("=IFERROR(1/0,#VALUE!)");
    expect(toLocalFormula("=_xlfn.XLOOKUP(1,A:A,B:B)")).toBe("=ÇAPRAZARA(1;A:A;B:B)");
    for (const [en, tr] of Object.entries(ERROR_TO_TR)) {
      const r = dongu(`=EĞERHATA(1;${tr})`);
      expect(r.kanonik, tr).toBe(`=IFERROR(1,${en})`);
      expect(r.geri, en).toBe(`=EĞERHATA(1;${tr})`);
      expect(fromLocalFormula(`=EĞERHATA(1;${trAlt(tr)})`), trAlt(tr)).toBe(`=IFERROR(1,${en})`);
    }
  });
});

describe("Türkçe işlev adları: fast-check", () => {
  const harfBuyuklugu = (s: string) => fc.array(fc.boolean(), { minLength: [...s].length, maxLength: [...s].length }).map((m) => [...s].map((c, i) => (m[i] ? trUst(c) : trAlt(c))).join(""));
  const islev = fc.constantFrom(...DEFS);

  it("rastgele harf büyüklüğüyle yazılan Türkçe ve İngilizce adlar doğru işleve çözülür (1000 yineleme)", () => {
    fc.assert(
      fc.property(
        islev.chain((d) => fc.tuple(fc.constant(d), harfBuyuklugu(d.tr), fc.array(fc.boolean(), { minLength: d.name.length, maxLength: d.name.length }))),
        ([d, trAd, maske]) => {
          const enAd = [...d.name].map((c, i) => (maske[i] ? c.toLowerCase() : c)).join("");
          expect(fromLocalFormula(`=${trAd}(A1)`, "tr")).toBe(`=${d.name}(A1)`);
          expect(fromLocalFormula(`=${enAd}(A1)`, "tr")).toBe(`=${d.name}(A1)`);
          expect(fromLocalFormula(`=${enAd}(A1)`, "en")).toBe(`=${d.name}(A1)`);
        },
      ),
      { numRuns: 1000 },
    );
  });

  it("rastgele iç içe Türkçe formüller tr → en → XLSX → tr aynen döner (1000 yineleme)", () => {
    const harf = fc.constantFrom(..."aıibcçgğİIoöşsuüzAÇĞÖŞÜ ;,.!?#()".split(""), '""');
    const metin = fc.array(harf, { maxLength: 8 }).map((p) => `"${p.join("")}"`);
    const sayi = fc.tuple(fc.integer({ min: 0, max: 99999 }), fc.option(fc.integer({ min: 0, max: 999 }), { nil: undefined })).map(([t, k]) => (k === undefined ? String(t) : `${t},${k}`));
    const hucre = fc.tuple(fc.constantFrom("A", "B", "Z", "AB", "$C"), fc.integer({ min: 1, max: 9999 })).map(([s, r]) => `${s}${r}`);
    const yaprak = fc.oneof(sayi, metin, hucre, fc.constantFrom("DOĞRU", "YANLIŞ", ...Object.values(ERROR_TO_TR)));
    const { ifade } = fc.letrec<{ ifade: string }>((tie) => ({
      ifade: fc.oneof(
        { depthSize: "small", withCrossShrink: true },
        yaprak,
        fc.tuple(islev, fc.array(tie("ifade"), { maxLength: 3 })).map(([d, a]) => `${d.tr}(${a.join(";")})`),
      ),
    }));
    fc.assert(
      fc.property(ifade, (govde) => {
        const tr = "=" + govde;
        const r = dongu(tr);
        expect(r.okunan).toBe(r.kanonik);
        expect(r.geri).toBe(tr);
        expect(sabitlerDisi(r.dosya)).toMatch(/^[\x20-\x7e]*$/);
      }),
      { numRuns: 1000 },
    );
  });
});

describe("BÜYÜKHARF / KÜÇÜKHARF / YAZIM.DÜZENİ: İ, ı, i, I", () => {
  /** Tek motor: A1'e metin yazılır, B1..B3 Türkçe formüllerle hesaplanır. */
  function motor(locale?: typeof EN_LOCALE) {
    const e = new FormulaEngine({ now: () => 0, random: () => 0.5, ...(locale ? { locale } : {}) });
    const s = e.addSheet("Sayfa1");
    const at = (a: string) => parseCellRef(a)!;
    for (const [a, f] of [["B1", "=BÜYÜKHARF(A1)"], ["B2", "=KÜÇÜKHARF(A1)"], ["B3", "=YAZIM.DÜZENİ(A1)"]]) e.setFormula(s, at(a).row, at(a).col, fromLocalFormula(f, "tr"));
    return (metin: string) => {
      e.setValue(s, 0, 0, metin);
      e.recalc();
      const g = (a: string) => e.getValue(s, at(a).row, at(a).col) as Scalar;
      return { ust: g("B1"), alt: g("B2"), duzen: g("B3") };
    };
  }

  it("Türkçe: i↔İ, ı↔I korunur", () => {
    const h = motor();
    expect(h("istanbul ılık iğne")).toEqual({ ust: "İSTANBUL ILIK İĞNE", alt: "istanbul ılık iğne", duzen: "İstanbul Ilık İğne" });
    expect(h("İSTANBUL IŞIK İĞDE")).toEqual({ ust: "İSTANBUL IŞIK İĞDE", alt: "istanbul ışık iğde", duzen: "İstanbul Işık İğde" });
    expect(h("iı İI")).toEqual({ ust: "İI İI", alt: "iı iı", duzen: "İı İı" });
  });

  it("İngilizce yerel ayarda Türkçe kural uygulanmaz", () => {
    const h = motor(EN_LOCALE);
    expect(h("india ink")).toEqual({ ust: "INDIA INK", alt: "india ink", duzen: "India Ink" });
  });

  it("rastgele Türkçe metinler: büyük/küçük harf dönüşümü noktalı/noktasız ayrımını korur (1000 yineleme)", () => {
    const h = motor();
    const metin = fc.array(fc.constantFrom(..."aıiIİbcçÇgğĞoöÖsşŞuüÜzZ 1-'.".split("")), { maxLength: 24 }).map((p) => p.join(""));
    fc.assert(
      fc.property(metin, (s) => {
        const r = h(s);
        expect(r.ust).toBe(trUst(s));
        expect(r.alt).toBe(trAlt(s));
        expect(String(r.ust)).not.toMatch(/[iı]/);
        expect(String(r.alt)).not.toMatch(/[Iİ]/);
        // Gidiş-dönüş: küçük → büyük → küçük aynı (İngilizce kuralda "ı" → "I" → "i" bozulurdu)
        expect(h(String(r.ust)).alt).toBe(trAlt(s));
        expect(h(String(r.alt)).ust).toBe(trUst(s));
        // YAZIM.DÜZENİ: harf dizisinin ilki büyük, kalanı küçük
        const beklenen = s.replace(/\p{L}+/gu, (w) => trUst(w[0]) + trAlt(w.slice(1)));
        expect(r.duzen).toBe(beklenen);
      }),
      { numRuns: 1000 },
    );
  });
});
