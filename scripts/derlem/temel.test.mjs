// Derlem / bulanık test araçlarının kendi testleri (node:test; bağımlılıksız).
// Kullanım: pnpm test:betikler   (node --test scripts/derlem/temel.test.mjs)
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { test } from "node:test";
import { crc32 } from "node:zlib";
import { aile, envanter, envanterFarki, envanterKaybi, iliskiTuru, kategori, uzantiOneki } from "./envanter.mjs";
import { kayipGruplari, kayipHaritasi, torba } from "./fark.mjs";
import { opcDenetle } from "./opc.mjs";
import { DERLEM, derlemDosyalari } from "./ortak.mjs";
import { bulanikKarsilastir, derlemKarsilastir, kipBelirle, opcKarsilastir, opcOzeti, poiKarsilastir, poiOzeti } from "./temel.mjs";

/** Sıkıştırmasız küçük ZIP (test paketleri için): girdiler [[ad, metin|Uint8Array]] sırasıyla yazılır. */
function zipYap(girdiler) {
  const enc = new TextEncoder();
  const yerel = [];
  const merkez = [];
  let ofset = 0;
  for (const [ad, icerik] of girdiler) {
    const a = enc.encode(ad);
    const v = typeof icerik === "string" ? enc.encode(icerik) : icerik;
    const crc = crc32(v);
    const b = Buffer.alloc(30 + a.length);
    b.writeUInt32LE(0x04034b50, 0);
    b.writeUInt16LE(20, 4);
    b.writeUInt32LE(crc, 14);
    b.writeUInt32LE(v.length, 18);
    b.writeUInt32LE(v.length, 22);
    b.writeUInt16LE(a.length, 26);
    b.set(a, 30);
    yerel.push(b, v);
    const m = Buffer.alloc(46 + a.length);
    m.writeUInt32LE(0x02014b50, 0);
    m.writeUInt16LE(20, 4);
    m.writeUInt16LE(20, 6);
    m.writeUInt32LE(crc, 16);
    m.writeUInt32LE(v.length, 20);
    m.writeUInt32LE(v.length, 24);
    m.writeUInt16LE(a.length, 28);
    m.writeUInt32LE(ofset, 42);
    m.set(a, 46);
    merkez.push(m);
    ofset += b.length + v.length;
  }
  const md = Buffer.concat(merkez);
  const son = Buffer.alloc(22);
  son.writeUInt32LE(0x06054b50, 0);
  son.writeUInt16LE(girdiler.length, 8);
  son.writeUInt16LE(girdiler.length, 10);
  son.writeUInt32LE(md.length, 12);
  son.writeUInt32LE(ofset, 16);
  return new Uint8Array(Buffer.concat([...yerel, md, son]));
}

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
/** Geçerli en küçük .docx paketi; `degistir` ile parçalar değiştirilir (null: parça çıkarılır). */
function docx(degistir = {}) {
  const p = {
    "[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="${CT_DOCX}"/></Types>`,
    "_rels/.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/document.xml": `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body><w:p><w:r><w:drawing><a:blip xmlns:a="x" r:embed="rId5"/></w:drawing></w:r></w:p><mc:AlternateContent><mc:Choice Requires="w14"><w14:checkbox/></mc:Choice><mc:Fallback/></mc:AlternateContent></w:body></w:document>`,
    "word/_rels/document.xml.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="${R}/image" Target="media/image1.png"/><Relationship Id="rId6" Type="${R}/hyperlink" Target="https://ornek.org" TargetMode="External"/><Relationship Id="rId7" Type="${R}/hyperlink" Target="#yer_imi"/></Relationships>`,
    "word/media/image1.png": "PNG",
    ...degistir,
  };
  return zipYap(Object.entries(p).filter(([, v]) => v !== null));
}
const turler = (h) => h.map((x) => x.tur).sort();

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

test("envanter: ilişki türleri ve uzantı ad alanı öğeleri; kayıp yalnız aynı ailede", () => {
  const e = envanter(docx());
  assert.deepEqual(e.iliskiler, { officeDocument: 1, image: 1, hyperlink: 2 });
  assert.deepEqual(e.ns, { "mc:AlternateContent": 1, "mc:Choice": 1, "mc:Fallback": 1, "w14:checkbox": 1 });
  assert.equal(iliskiTuru("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"), "metadata/core-properties");
  assert.equal(iliskiTuru("http://schemas.microsoft.com/office/2006/relationships/vbaProject"), "vbaProject");
  assert.equal(uzantiOneki("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"), "x14");
  assert.equal(uzantiOneki("http://schemas.microsoft.com/office/powerpoint/2099/main"), "ms:powerpoint/2099/main", "bilinmeyen ofis uzantısı yoldan adlanır");
  assert.equal(uzantiOneki("http://schemas.openxmlformats.org/wordprocessingml/2006/main"), null, "ana ad alanı uzantı değil");
  const smart = envanter(readFileSync(join(DERLEM, "poi/SmartArt.pptx")));
  assert.ok(smart.ns["dgm:pt"] > 0 && smart.iliskiler.diagramData === 1, "SmartArt ad alanı ve ilişkisi sayılır");
  const A = { aile: "ooxml", parcalar: {}, ogeler: {}, iliskiler: { vbaProject: 1, styles: 1 }, ns: { "x14:dataValidation": 3 } };
  const B = { aile: "ooxml", parcalar: {}, ogeler: {}, iliskiler: { styles: 1 }, ns: { "x14:dataValidation": 1 } };
  assert.deepEqual(envanterKaybi(A, B), { "ilişki:vbaProject": 1, "ns:x14:dataValidation": 2 });
  assert.deepEqual(envanterKaybi(A, { ...B, aile: "odf" }), {}, "farklı aile: ilişki ve ad alanı karşılaştırılmaz");
  assert.deepEqual(envanterFarki(A, B), { "ilişki:vbaProject": "1→0", "ns:x14:dataValidation": "3→1" });
  assert.deepEqual(kayipGruplari({ "parça:theme": 2, "ilişki:theme": 2, "ns:mc:Choice": 1, "öğe:yorum": 1, kelime: 4, "düğüm.paragraph": 1 }), { parça: 2, ilişki: 2, ns: 1, öğe: 1, model: 5 });
});

test("opc: geçerli paket temiz; bozulmalar türüne göre yakalanır", () => {
  assert.deepEqual(opcDenetle(docx(), { uzanti: "docx" }), [], "dış ve yalnız fragmanlı hedef hata değildir");
  assert.deepEqual(turler(opcDenetle(docx(), { uzanti: "docm" })), ["uzantı-içerik-türü"], ".docm uzantısı makrolu ana tür ister");
  assert.deepEqual(turler(opcDenetle(docx({ "word/media/image1.png": null }))), ["hedef-yok"], "hedefi olmayan ilişki");
  const kirik = docx({ "word/_rels/document.xml.rels": `<Relationships xmlns="x"><Relationship Id="rId9" Type="${R}/image" Target="media/image1.png"/></Relationships>` });
  assert.deepEqual(turler(opcDenetle(kirik)), ["kırık-rId"], "XML'deki r:embed .rels'te yok");
  const tur = docx({
    "[Content_Types].xml": `<Types xmlns="x"><Default Extension="rels" ContentType="a"/><Default Extension="xml" ContentType="b"/><Default Extension="XML" ContentType="b"/><Override PartName="/word/document.xml" ContentType="${CT_DOCX}"/><Override PartName="/word/yok.xml" ContentType="c"/></Types>`,
  });
  assert.deepEqual(turler(opcDenetle(tur)), ["fazla-override", "içerik-türü-yok", "yinelenen-içerik-türü"], "png türü yok, fazla Override, yinelenen Default");
  const yinelenen = turler(opcDenetle(zipYap([["[Content_Types].xml", "<Types/>"], ["a.xml", "1"], ["A.xml", "2"]])));
  assert.ok(yinelenen.includes("yinelenen-parça"), "yalnız harf büyüklüğü farklı ad yinelenen parçadır");
  assert.ok(yinelenen.includes("ana-belge-yok"));
  const sahipsiz = docx({ "word/_rels/yok.xml.rels": `<Relationships xmlns="x"><Relationship Id="rId1" Type="t" Target="document.xml"/><Relationship Id="rId1" Type="t" Target="document.xml"/></Relationships>` });
  assert.deepEqual(turler(opcDenetle(sahipsiz)), ["sahipsiz-rels", "yinelenen-rId"]);
  assert.deepEqual(turler(opcDenetle(new Uint8Array([1, 2, 3]))), ["zip-okunamadı"]);
});

test("opc: ODF mimetype ve manifest", () => {
  const MT = "application/vnd.oasis.opendocument.text";
  const man = (ek = "") =>
    `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="${MT}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>${ek}</manifest:manifest>`;
  assert.deepEqual(opcDenetle(zipYap([["mimetype", MT], ["content.xml", "<x/>"], ["META-INF/manifest.xml", man()]]), { uzanti: "odt" }), []);
  const bozuk = zipYap([
    ["content.xml", "<x/>"],
    ["mimetype", MT],
    ["styles.xml", "<x/>"],
    ["META-INF/manifest.xml", man('<manifest:file-entry manifest:full-path="Pictures/yok.png" manifest:media-type="image/png"/>')],
  ]);
  assert.deepEqual(turler(opcDenetle(bozuk, { uzanti: "ods" })), ["odf-manifest-eksik", "odf-manifest-fazla", "odf-mimetype", "uzantı-içerik-türü"]);
});

test("opc temel: hata artışı her kipte HATA, azalma iyileşme", () => {
  const satirlar = [
    { anahtar: "a.docx →docx", hatalar: [{ tur: "hedef-yok" }, { tur: "hedef-yok" }] },
    { anahtar: "b.xlsx →xlsx", hatalar: [] },
  ];
  const oz = opcOzeti(satirlar);
  assert.deepEqual(oz, { "a.docx →docx": { "hedef-yok": 2 } }, "yalnız hatalı çıktılar saklanır");
  const temel = { "a.docx →docx": { "hedef-yok": 1 }, "b.xlsx →xlsx": { "kırık-rId": 1 }, "c.pptx →pptx": { "hedef-yok": 1 } };
  const b = opcKarsilastir(oz, temel, satirlar.map((s) => s.anahtar), "baslangic");
  assert.deepEqual(b.hatalar, ["a.docx →docx: yeni paket hatası: hedef-yok 1→2"]);
  assert.deepEqual(b.iyilesmeler, ["b.xlsx →xlsx: paket hatası azaldı: kırık-rId 1→0"]);
  assert.equal(b.bilgiler.length, 1, "bu koşuda bakılmayan çıktı bilgi");
});

test("POI kapısı: parça kaybı artışı ve açılamayan dosya her kipte HATA", () => {
  const YEREL = { belge: "docx", tablo: "xlsx", sunum: "pptx" };
  const dosya = (anahtar, app, kayip, acildi = true) => ({ anahtar, app, acildi, turlar: acildi ? [{ bicim: YEREL[app], kayip }, { bicim: "odt", kayip: { "parça:x": 99 } }] : undefined });
  const oz = poiOzeti([
    dosya("test-derlem/poi/a.docx", "belge", { "parça:theme": 1, "ilişki:theme": 1, kelime: 2 }),
    dosya("test-derlem/poi/b.docx", "belge", {}, false),
    dosya("test-derlem/lo/c.docx", "belge", { "parça:theme": 5 }),
  ]);
  assert.deepEqual(oz, { belge: { dosya: 2, acildi: 1, parça: 1, ilişki: 1, ns: 0, öğe: 0, model: 2 } }, "yalnız POI dosyaları ve yerel biçim turu");
  const temel = { belge: { dosya: 2, acildi: 2, parça: 0, ilişki: 3, ns: 0, öğe: 0, model: 2 } };
  const b = poiKarsilastir(oz, temel, "baslangic");
  assert.deepEqual(b.hatalar, ["belge: POI derleminde açılan dosya 2→1", "belge: POI parça kaybı 0→1"]);
  assert.deepEqual(b.iyilesmeler, ["belge: POI ilişki kaybı 3→1"]);
  const g = poiKarsilastir({ belge: { ...temel.belge, ns: 4 } }, temel, "baslangic");
  assert.deepEqual([g.hatalar.length, g.gerilemeler.length], [0, 1], "ad alanı kaybı artışı kipe bağlı gerileme");
});

test("derlem dosya listesi: depodaki alt küme ve fikstürler, anahtarlar benzersiz", () => {
  const d = derlemDosyalari();
  assert.ok(d.filter((x) => x.kaynak === "derlem").length >= 24);
  const poi = d.filter((x) => x.anahtar.startsWith("test-derlem/poi/"));
  assert.ok(poi.length >= 40, "POI derlemi tura girer");
  assert.deepEqual(new Set(poi.map((x) => x.app)), new Set(["belge", "tablo", "sunum"]));
  assert.ok(d.some((x) => x.anahtar === "apps/tablo/src/io/__fixtures__/ornek.ods"));
  assert.equal(new Set(d.map((x) => x.anahtar)).size, d.length);
});
