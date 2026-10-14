// Bozuk/hatalı dosya örnekleri üretir (bulanık test girdileri). Tohumlu ve deterministiktir: aynı
// tohum dosyaları ve aynı betikle her seferinde aynı 756 örnek çıkar; örnekler depoya konmaz.
//
// Tohumlar depodaki derlemden (test-derlem/) ve uygulama fikstürlerinden alınır. Mutasyonlar:
// bayt bozma, kesik dosya, başlık bozma, ZIP parçası silme, ana XML'i kesme/boşaltma/çöp, sayısal
// öznitelikleri uç değerlere çekme, derin iç içe yapı, varlık genişlemesi (billion laughs), zip bombası
// (300 MB), ilişki hedefi yok / döngü, uç hücre başvurusu, CFB FAT döngüsü/halkası, RTF/Markdown
// derinliği, yanlış uzantı, boş/sahte ZIP ve CFB, CSV uç durumları.
//
// Kullanım: node scripts/bulanik/uret.mjs [--zorla]   → <CIKTI>/bulanik/<uygulama>/NNN_ad + liste.json
// (bulanik/calistir.mjs gerekirse kendisi çağırır; tohumlar ve betik değişmediyse yeniden üretmez.)
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { basename, extname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { argumanlar, CIKTI, DERLEM, GORSEL, KOK } from "../derlem/ortak.mjs";

export const BULANIK = join(CIKTI, "bulanik");
const D = join(DERLEM, "lo");
const P = join(DERLEM, "py");
export const TOHUMLAR = {
  belge: [`${D}/lo-rapor.docx`, `${D}/lo-rapor.odt`, `${D}/lo-rapor.rtf`, `${D}/lo-rapor.doc`, `${P}/py-belge.docx`, join(KOK, "apps/belge/src/io/__fixtures__/bicimler/ornek.md")],
  tablo: [`${D}/lo-butce.xlsx`, `${D}/lo-butce.ods`, `${D}/lo-butce.xls`, `${P}/py-tablo.xlsx`, join(KOK, "apps/tablo/test/fixtures/openpyxl/karmasik.xlsx")],
  sunum: [`${D}/lo-sunu.pptx`, `${D}/lo-sunu.odp`, `${D}/lo-sunu.ppt`, `${P}/py-sunu.pptx`],
};
const GENEL = {
  docx: `${D}/lo-rapor.docx`, xlsx: `${D}/lo-butce.xlsx`, pptx: `${D}/lo-sunu.pptx`, odt: `${D}/lo-rapor.odt`, ods: `${D}/lo-butce.ods`, odp: `${D}/lo-sunu.odp`,
  doc: `${D}/lo-rapor.doc`, xls: `${D}/lo-butce.xls`, ppt: `${D}/lo-sunu.ppt`, png: join(GORSEL, "grafik.png"), html: `${D}/lo-rapor.html`, rtf: `${D}/lo-rapor.rtf`,
};

/** Tohumların ve bu betiğin özeti: değişmediyse örnekler yeniden üretilmez. */
export function tohumOzeti() {
  const h = createHash("sha256");
  h.update(readFileSync(fileURLToPath(import.meta.url)));
  for (const f of [...Object.values(TOHUMLAR).flat(), ...Object.values(GENEL)]) h.update(readFileSync(f));
  return h.digest("hex").slice(0, 16);
}

export async function uret({ zorla = false } = {}) {
  const ozet = tohumOzeti();
  const listeYolu = join(BULANIK, "liste.json");
  if (!zorla && existsSync(listeYolu)) {
    const eski = JSON.parse(readFileSync(listeYolu, "utf8"));
    if (eski.ozet === ozet) return eski;
  }
  const t0 = performance.now();
  // JSZip uygulamaların zaten kullandığı bağımlılıktır (yeni bağımlılık yok).
  const JSZip = createRequire(join(KOK, "apps/belge/package.json"))("jszip");
  rmSync(BULANIK, { recursive: true, force: true });
  let tohum = 20261014;
  const rnd = () => (tohum = (tohum * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
  const ri = (n) => Math.floor(rnd() * n);
  // ZIP girdilerinin tarihi sabit: çıktı bayt bayt aynı olsun.
  const TARIH = new Date(Date.UTC(2026, 0, 1));

  const liste = [];
  const kaydet = (app, ad, bytes, tur, tohumDosyasi) => {
    mkdirSync(join(BULANIK, app), { recursive: true });
    const yol = join(BULANIK, app, `${String(liste.length).padStart(3, "0")}_${ad}`);
    writeFileSync(yol, bytes);
    liste.push({ app, ad, yol, tur, tohum: tohumDosyasi ? basename(tohumDosyasi) : "-" });
  };
  const zipMi = (b) => b[0] === 0x50 && b[1] === 0x4b;
  const cfbMi = (b) => b[0] === 0xd0 && b[1] === 0xcf;
  const cevir = (b, n) => {
    const c = new Uint8Array(b);
    for (let i = 0; i < n; i++) c[ri(c.length)] ^= 1 << ri(8);
    return c;
  };
  const zipDegistir = async (b, fn) => {
    const z = await JSZip.loadAsync(b);
    await fn(z);
    return z.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  };
  const ANA = { docx: "word/document.xml", xlsx: "xl/worksheets/sheet1.xml", pptx: "ppt/slides/slide1.xml", odt: "content.xml", ods: "content.xml", odp: "content.xml" };
  const enc = (s) => new TextEncoder().encode(s);

  for (const [app, tohumlar] of Object.entries(TOHUMLAR)) {
    for (const sf of tohumlar) {
      const b = new Uint8Array(readFileSync(sf));
      const ext = extname(sf).slice(1);
      const kok = basename(sf, "." + ext);
      // 1) rastgele bayt bozma
      for (const n of [1, 8, 64, 512, 4096]) for (let k = 0; k < 2; k++) kaydet(app, `${kok}_flip${n}_${k}.${ext}`, cevir(b, n), `bayt-bozma-${n}`, sf);
      // 2) kesik dosya
      for (const yuzde of [1, 10, 50, 90, 99]) kaydet(app, `${kok}_kesik${yuzde}.${ext}`, b.slice(0, Math.floor((b.length * yuzde) / 100)), `kesik-%${yuzde}`, sf);
      // 3) başlık bozma (ilk 64 bayt)
      {
        const c = new Uint8Array(b);
        for (let i = 0; i < 64; i++) c[i] = ri(256);
        kaydet(app, `${kok}_baslik.${ext}`, c, "başlık-bozma", sf);
      }
      if (zipMi(b)) {
        const z = await JSZip.loadAsync(b);
        const adlar = Object.keys(z.files).filter((n) => !z.files[n].dir);
        // 4) önemli parçaları tek tek silme
        const onemli = adlar.filter((n) => /Content_Types|document\.xml$|workbook\.xml$|presentation\.xml$|styles\.xml$|sharedStrings|_rels\/|sheet1\.xml|slide1\.xml|slideLayout1\.xml|slideMaster1\.xml|theme1|content\.xml|manifest|mimetype|numbering|footnotes|comments|header1|chart1|drawing1|table1/.test(n));
        for (const n of onemli.slice(0, 16)) kaydet(app, `${kok}_sil_${n.replace(/[/[\]]/g, "_")}.${ext}`, await zipDegistir(b, async (zz) => zz.remove(n)), "zip-parça-silme", sf);
        // 5) ana XML'i yarıda kesme / boşaltma / çöp
        const ana = ANA[ext] ?? adlar.find((n) => n.endsWith(".xml"));
        const xml = await z.file(ana).async("string");
        const yaz = (zz, ad, veri) => zz.file(ad, veri, { date: TARIH, createFolders: false });
        kaydet(app, `${kok}_xmlkesik.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, ana, xml.slice(0, xml.length >> 1))), "zip-xml-yarıda-kesik", sf);
        kaydet(app, `${kok}_xmlbos.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, ana, "")), "zip-xml-boş", sf);
        kaydet(app, `${kok}_xmlcop.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, ana, cevir(enc(xml), 200))), "zip-xml-çöp", sf);
        // 6) sayısal öznitelikleri uç değerlere çekme (mantık hataları, dev bellek ayırma)
        for (const [etiket, deger] of [["dev", "2147483647"], ["eksi", "-1"], ["nan", "NaN"], ["sifir", "0"]]) {
          const mut = xml.replace(
            /(\s(?:w:val|w:w|w:gridSpan|gridSpan|cx|cy|x|y|sz|r|ref|count|spans|ht|width|s|t|w:sz|table:number-columns-repeated|table:number-rows-repeated|table:number-columns-spanned|table:number-rows-spanned|svg:width|svg:height|w:fill|idx|lvl|w:ilvl|w:numId)=")([^"]*)(")/g,
            (m, a, v, c) => (rnd() < 0.5 && /^-?\d/.test(v) ? a + deger + c : m),
          );
          kaydet(app, `${kok}_oznitelik_${etiket}.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, ana, mut)), `öznitelik-${etiket}`, sf);
        }
        // 7) derin iç içe yapı (özyineleme taşması)
        const derin = xml.replace(/(<\/[a-zA-Z:]*body>|<\/office:text>|<\/sheetData>|<\/p:spTree>|<\/table:table>|<\/office:presentation>)/, (m) => "<x:d>".repeat(20000) + "</x:d>".repeat(20000) + m);
        kaydet(app, `${kok}_derin.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, ana, derin)), "derin-iç-içe", sf);
        // 8) varlık genişlemesi (billion laughs)
        const lol =
          '<?xml version="1.0"?><!DOCTYPE lolz [<!ENTITY lol "lollollollollollollollollollol"><!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;"><!ENTITY lol3 "&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;"><!ENTITY lol4 "&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;"><!ENTITY lol5 "&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;"><!ENTITY lol6 "&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;"><!ENTITY lol7 "&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;&lol6;">]>';
        kaydet(app, `${kok}_varlik.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, ana, xml.replace(/^<\?xml[^>]*\?>/, lol).replace(/(<w:t[^>]*>|<a:t>|<text:p[^>]*>|<v>)/, "$1&lol7;"))), "varlık-genişlemesi", sf);
        // 9) zip bombası: ana XML'e 300 MB boşluk (sıkıştırılmış ~300 KB)
        kaydet(app, `${kok}_bomba.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, ana, xml.replace(/>\s*$/, ">" + " ".repeat(300 * 1024 * 1024 - xml.length)))), "zip-bombası-300MB", sf);
        // 10) ilişki hedefi yok / döngüsel ilişki
        const relAd = adlar.find((n) => /_rels\/(document|workbook|presentation)\.xml\.rels$/.test(n));
        if (relAd) {
          const rel = await z.file(relAd).async("string");
          kaydet(app, `${kok}_rels_yok.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, relAd, rel.replace(/Target="[^"]+"/g, 'Target="../../yok/olmayan.xml"'))), "ilişki-hedefi-yok", sf);
          kaydet(app, `${kok}_rels_dongu.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, relAd, rel.replace(/Target="([^"]+)"/g, (m) => (rnd() < 0.5 ? `Target="${basename(relAd, ".rels")}"` : m)))), "ilişki-döngüsü", sf);
        }
        // 11) uç hücre başvurusu / dev boyut / satır 0 / döngüsel formül (xlsx)
        if (ext === "xlsx") {
          kaydet(app, `${kok}_xfd.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, ana, xml.replace(/<dimension ref="[^"]*"\/>/, '<dimension ref="A1:XFD1048576"/>').replace("</sheetData>", '<row r="1048576"><c r="XFD1048576"><v>1</v></c></row></sheetData>'))), "uç-hücre-XFD1048576", sf);
          kaydet(app, `${kok}_birlesik_tum.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, ana, xml.replace(/<mergeCells[^>]*>/, '<mergeCells count="1"><mergeCell ref="A1:XFD1048576"/>').replace(/<\/sheetData>(?![\s\S]*<mergeCells)/, '</sheetData><mergeCells count="1"><mergeCell ref="A1:XFD1048576"/></mergeCells>'))), "tüm-sayfa-birleşik", sf);
          kaydet(app, `${kok}_satir_sifir.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, ana, xml.replace(/<row r="\d+"/, '<row r="0"').replace(/<c r="[A-Z]+\d+"/, '<c r="A0"'))), "satır-0", sf);
          kaydet(app, `${kok}_formul_dongu.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, ana, xml.replace("</sheetData>", '<row r="2000"><c r="A2000"><f>B2000+1</f></c><c r="B2000"><f>A2000+1</f></c><c r="C2000"><f>' + "SUM(".repeat(2000) + "1" + ")".repeat(2000) + "</f></c></row></sheetData>"))), "döngüsel/derin formül", sf);
        }
        if (ext === "docx") {
          kaydet(app, `${kok}_gridspan.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, ana, xml.replace(/<w:tcPr>/g, '<w:tcPr><w:gridSpan w:val="100000"/>'))), "dev-gridSpan", sf);
          kaydet(app, `${kok}_numid.${ext}`, await zipDegistir(b, async (zz) => yaz(zz, ana, xml.replace(/<w:pPr>/, '<w:pPr><w:numPr><w:ilvl w:val="99"/><w:numId w:val="99999"/></w:numPr>'))), "olmayan-numara", sf);
        }
        if (ext === "pptx") {
          kaydet(app, `${kok}_sldsz.${ext}`, await zipDegistir(b, async (zz) => { const pz = await zz.file("ppt/presentation.xml").async("string"); yaz(zz, "ppt/presentation.xml", pz.replace(/<p:sldSz[^>]*\/>/, '<p:sldSz cx="0" cy="-5"/>')); }), "slayt-boyutu-0", sf);
        }
      }
      if (cfbMi(b)) {
        // 12) CFB: FAT zincirinde döngü (dizin sektörü kendine işaret eder) → sonsuz döngü adayı
        const dv = new DataView(b.buffer.slice(0));
        const c = new Uint8Array(dv.buffer);
        const ss = 1 << dv.getUint16(0x1e, true);
        const dizinBas = dv.getUint32(0x30, true);
        const fatOff = (dv.getUint32(0x4c, true) + 1) * ss;
        const c1 = new Uint8Array(c);
        new DataView(c1.buffer).setUint32(fatOff + dizinBas * 4, dizinBas, true);
        kaydet(app, `${kok}_fat_dongu.${ext}`, c1, "cfb-fat-döngüsü", sf);
        const c2 = new Uint8Array(c);
        for (let i = 0; i < ss / 4; i++) new DataView(c2.buffer).setUint32(fatOff + i * 4, (i + 1) % (ss / 4), true);
        kaydet(app, `${kok}_fat_halka.${ext}`, c2, "cfb-fat-halka", sf);
        const c3 = new Uint8Array(c);
        new DataView(c3.buffer).setUint16(0x1e, 30, true);
        kaydet(app, `${kok}_sektor_dev.${ext}`, c3, "cfb-sektör-boyu-2^30", sf);
        const c4 = new Uint8Array(c);
        new DataView(c4.buffer).setUint32(0x2c, 0x7fffffff, true);
        kaydet(app, `${kok}_fat_sayisi.${ext}`, c4, "cfb-fat-sayısı-dev", sf);
        const c5 = new Uint8Array(c);
        new DataView(c5.buffer).setUint32(0x30, 0xfffffffe, true);
        kaydet(app, `${kok}_dizin_yok.${ext}`, c5, "cfb-dizin-yok", sf);
        // dizin girdisinde akış boyutunu devleştir
        const c6 = new Uint8Array(c);
        const dizinOff = (dizinBas + 1) * ss;
        for (let e = 1; e < 4; e++) new DataView(c6.buffer).setUint32(dizinOff + e * 128 + 0x78, 0x7ffffff0, true);
        kaydet(app, `${kok}_akis_dev.${ext}`, c6, "cfb-akış-boyu-dev", sf);
      }
      if (ext === "rtf") {
        const s = new TextDecoder("latin1").decode(b);
        kaydet(app, `${kok}_derin.${ext}`, enc("{\\rtf1 " + "{".repeat(200000) + "x" + "}".repeat(200000) + "}"), "rtf-derin-grup", sf);
        kaydet(app, `${kok}_dengesiz.${ext}`, enc(s.replace(/\}/g, "")), "rtf-kapanmayan-grup", sf);
        kaydet(app, `${kok}_bin.${ext}`, enc(s.replace("\\par", "\\bin2147483647 \\par")), "rtf-dev-bin", sf);
        kaydet(app, `${kok}_unicode.${ext}`, enc(s.replace("\\par", "\\u-99999? \\uc999999 \\u4294967295 x\\par")), "rtf-uç-unicode", sf);
        kaydet(app, `${kok}_pict.${ext}`, enc(s.replace("\\par", "{\\pict\\pngblip\\picw-5\\pich999999999 zzzz}\\par")), "rtf-bozuk-resim", sf);
      }
      if (ext === "md") {
        const s = new TextDecoder().decode(b);
        kaydet(app, `${kok}_derin.md`, enc(">".repeat(50000) + " x\n" + "- ".repeat(20000) + "y\n" + "[".repeat(50000) + "\n"), "md-derin", sf);
        kaydet(app, `${kok}_tablo.md`, enc("|" + "a|".repeat(20000) + "\n|" + "-|".repeat(20000) + "\n" + s), "md-dev-tablo", sf);
      }
    }
  }
  // 13) genel: 0 bayt, rastgele çöp, yanlış uzantı (başka biçimin içeriği), sahte imzalar
  const herhangi = Object.fromEntries(Object.entries(GENEL).map(([k, v]) => [k, readFileSync(v)]));
  const pdfTaslak = enc("%PDF-1.4\n1 0 obj<</Type/Catalog>>endobj\ntrailer<</Root 1 0 R>>\n%%EOF");
  const uzantilar = { belge: ["docx", "doc", "odt", "rtf", "md", "html", "txt"], tablo: ["xlsx", "xls", "ods", "csv"], sunum: ["pptx", "ppt", "odp"] };
  for (const [app, es] of Object.entries(uzantilar)) {
    for (const e of es) {
      kaydet(app, `bos.${e}`, new Uint8Array(0), "0-bayt");
      kaydet(app, `cop.${e}`, Uint8Array.from({ length: 4096 }, () => ri(256)), "rastgele-çöp");
      kaydet(app, `pdf.${e}`, pdfTaslak, "yanlış-uzantı:pdf");
      kaydet(app, `png.${e}`, herhangi.png, "yanlış-uzantı:png");
      for (const [k, v] of Object.entries(herhangi)) if (k !== e && k !== "png") kaydet(app, `${k}-icerik.${e}`, v, `yanlış-uzantı:${k}`);
      kaydet(app, `pk_cop.${e}`, Uint8Array.from([0x50, 0x4b, 0x03, 0x04, ...Array.from({ length: 2000 }, () => ri(256))]), "sahte-zip-imzası");
      kaydet(app, `cfb_cop.${e}`, Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, ...Array.from({ length: 4000 }, () => ri(256))]), "sahte-cfb-imzası");
    }
  }
  // boş ZIP ve yalnız [Content_Types].xml içeren ZIP
  {
    const z = new JSZip();
    const bos = await z.generateAsync({ type: "uint8array" });
    z.file("[Content_Types].xml", "<Types/>", { date: TARIH, createFolders: false });
    const ct = await z.generateAsync({ type: "uint8array" });
    for (const [app, e] of [["belge", "docx"], ["tablo", "xlsx"], ["sunum", "pptx"]]) {
      kaydet(app, `bos_zip.${e}`, bos, "boş-zip");
      kaydet(app, `yalniz_ct.${e}`, ct, "yalnız-Content_Types");
    }
  }
  // CSV özel: kapanmayan tırnak, dev satır, çok satır, NUL baytları, UTF-16, formül enjeksiyonu
  const csvler = {
    "kapanmayan_tirnak.csv": '"a;b\n' + "x;".repeat(200000),
    "dev_satir.csv": "a;".repeat(300000) + "\n1;2\n",
    "cok_satir.csv": "a;b;c\n".repeat(400000),
    "nul.csv": "a;b\0c\n\0\0\0\n",
    "utf16.csv": "﻿a;b\n1;2\n",
    "formul_enj.csv": '=cmd|"/C calc"!A0;=HYPERLINK("http://x");@SUM(1)\n',
  };
  for (const [n, s] of Object.entries(csvler)) kaydet("tablo", n, n === "utf16.csv" ? new Uint8Array(Buffer.from(s, "utf16le")) : enc(s), "csv-" + n.replace(".csv", ""));

  const sayilar = {};
  for (const x of liste) sayilar[x.app] = (sayilar[x.app] ?? 0) + 1;
  const sonuc = { ozet, tohum: 20261014, sayilar, ornekler: liste.map((x) => ({ ...x, yol: x.yol.slice(BULANIK.length + 1) })) };
  writeFileSync(listeYolu, JSON.stringify(sonuc, null, 1));
  console.log(`bulanık örnekler üretildi: ${liste.length} (${Object.entries(sayilar).map(([k, v]) => `${k} ${v}`).join(", ")}) ${Math.round(performance.now() - t0)} ms`);
  return sonuc;
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) await uret({ zorla: !!argumanlar().zorla });
