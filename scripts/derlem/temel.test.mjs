// Derlem / bulanık test araçlarının kendi testleri (node:test; bağımlılıksız).
// Kullanım: pnpm test:betikler   (node --test scripts/derlem/temel.test.mjs)
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { test } from "node:test";
import { aile, envanter, envanterFarki, envanterKaybi, kategori } from "./envanter.mjs";
import { kayipHaritasi, torba } from "./fark.mjs";
import { DERLEM, derlemDosyalari } from "./ortak.mjs";
import { bulanikKarsilastir, derlemKarsilastir, kipBelirle } from "./temel.mjs";

test("envanter: LibreOffice .docx ve .odt anlamlı öğeleri sayar", () => {
  const docx = envanter(readFileSync(join(DERLEM, "lo/lo-rapor.docx")));
  assert.equal(docx.aile, "ooxml");
  for (const k of ["paragraf", "tablo", "resim", "dipnot", "yorum", "üstBilgi", "altBilgi"]) assert.ok(docx.ogeler[k] > 0, `docx ${k}`);
  assert.ok(docx.parcalar["ana-belge"] === 1, "ana belge parçası içerik türünden tanınır");
  const odt = envanter(readFileSync(join(DERLEM, "lo/lo-rapor.odt")));
  assert.equal(odt.aile, "odf");
  assert.equal(odt.ogeler.dipnot, docx.ogeler.dipnot, "aynı belgenin iki biçiminde dipnot sayısı eşit");
  assert.equal(odt.ogeler.tablo, docx.ogeler.tablo);
  assert.equal(envanter(readFileSync(join(DERLEM, "lo/lo-rapor.rtf"))), null, "ZIP olmayan dosyada envanter yok");
});

test("envanter: kayıp yalnız azalmadır; parça karşılaştırması aynı ailede", () => {
  const A = { aile: "ooxml", parcalar: { "ana-belge": 1, "wordprocessingml.comments": 1, medya: 3 }, ogeler: { yorum: 2, paragraf: 10 } };
  const B = { aile: "ooxml", parcalar: { "ana-belge": 1, medya: 1 }, ogeler: { yorum: 0, paragraf: 12 } };
  assert.deepEqual(envanterKaybi(A, B), { "parça:wordprocessingml.comments": 1, "parça:medya": 2, "öğe:yorum": 2 });
  assert.deepEqual(envanterKaybi(A, { ...B, aile: "odf" }), { "öğe:yorum": 2 }, "farklı aile: yalnız öğeler");
  assert.deepEqual(envanterFarki(B, B), {});
  assert.equal(kategori("word/media/image3.png"), "medya");
  assert.equal(kategori("ppt/slides/slide12.xml"), "ppt/slides/slide#.xml");
  assert.equal(aile(null), null);
});

test("fark: kelime torbası ve kayıp haritası", () => {
  assert.deepEqual(torba("Bir iki İki üç", "bir iki dört"), { top: 4, kayip: 2, eklenen: 1 });
  const k = kayipHaritasi("sunum", { kelime: { kayip: 3 }, slaytSayisi: [5, 4], sekil: { sp: [10, 8], pic: [1, 2] }, bicim: { renk: [56, 250] } }, {}, null);
  assert.deepEqual(k, { slayt: 1, kelime: 3, "şekil.sp": 2 }, "artışlar (pic, renk) kayıp sayılmaz");
});

test("temel: ağır durum her kipte hata, kayıp artışı kipe göre", () => {
  const temel = {
    "a.docx": { durum: "tamam", acildi: true, turlar: { docx: { kayip: { kelime: 2 } } } },
    "b.xlsx": { durum: "tamam", acildi: true, turlar: { xlsx: {} } },
    "c.pptx": { durum: "zaman-asimi", acildi: false },
  };
  const simdi = {
    "a.docx": { durum: "tamam", acildi: true, turlar: { docx: { kayip: { kelime: 5, "öğe:yorum": 1 } } } },
    "b.xlsx": { durum: "zaman-asimi", acildi: false },
    "c.pptx": { durum: "zaman-asimi", acildi: false },
    "d.odt": { durum: "tamam", acildi: true, yigin: true },
  };
  const bas = derlemKarsilastir(simdi, temel, "baslangic");
  assert.equal(bas.hatalar.length, 2, "yeni takılma (b) + yığın taşması (d); bilinen takılma (c) hata değil");
  assert.ok(bas.hatalar.some((s) => s.startsWith("b.xlsx")) && bas.hatalar.some((s) => s.startsWith("d.odt")));
  assert.equal(bas.gerilemeler.length, 1, "yeni kayıp başlangıç kipinde yalnız rapor");
  assert.match(bas.gerilemeler[0], /kelime 2→5, öğe:yorum 0→1/);
  const siki = derlemKarsilastir(simdi, temel, "siki");
  assert.equal(siki.hatalar.length, 3, "sıkı kipte gerileme de hata");
  const iyi = derlemKarsilastir({ "a.docx": { durum: "tamam", acildi: true, turlar: { docx: {} } } }, temel, "siki");
  assert.equal(iyi.hatalar.length, 0);
  assert.equal(iyi.iyilesmeler.length, 1);
});

test("temel: bulanık sınıf değişimleri", () => {
  const T = { ornekler: { "tablo/x.xlsx": "anlaşılır hata", "tablo/y.xlsx": "takılma", "belge/z.docx": "ham hata", "belge/w.docx": "açıldı" } };
  const S = { ornekler: { "tablo/x.xlsx": "ham hata", "tablo/y.xlsx": "takılma", "belge/z.docx": "anlaşılır hata", "belge/w.docx": "yığın taşması", "sunum/yeni.pptx": "takılma" } };
  const b = bulanikKarsilastir(S, T, "baslangic");
  assert.deepEqual(b.hatalar.map((s) => s.split(":")[0]).sort(), ["belge/w.docx", "sunum/yeni.pptx"]);
  assert.deepEqual(b.gerilemeler, ["tablo/x.xlsx: anlaşılır hata → ham hata"]);
  assert.deepEqual(b.iyilesmeler, ["belge/z.docx: ham hata → anlaşılır hata"]);
});

test("kip: boş ortam değişkeni temel.json'daki kipi ezmez", () => {
  const eski = process.env.DERLEM_KIP;
  process.env.DERLEM_KIP = "";
  assert.equal(kipBelirle({}, { kip: "siki" }), "siki");
  process.env.DERLEM_KIP = "baslangic";
  assert.equal(kipBelirle({}, { kip: "siki" }), "baslangic");
  assert.equal(kipBelirle({ kip: "siki" }, { kip: "baslangic" }), "siki");
  if (eski === undefined) delete process.env.DERLEM_KIP;
  else process.env.DERLEM_KIP = eski;
});

test("derlem dosya listesi: depodaki alt küme ve fikstürler, anahtarlar benzersiz", () => {
  const d = derlemDosyalari();
  assert.ok(d.filter((x) => x.kaynak === "derlem").length >= 24);
  assert.ok(d.some((x) => x.anahtar === "apps/tablo/src/io/__fixtures__/ornek.ods"));
  assert.equal(new Set(d.map((x) => x.anahtar)).size, d.length);
});
