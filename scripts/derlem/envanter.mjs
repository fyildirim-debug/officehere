// OOXML / ODF paket envanteri: ZIP parçaları ve anlamlı öğe sayıları.
//
// Kaydet-aç turunda model düzeyindeki karşılaştırmaya ek olarak dosyanın kendisine bakılır:
//  - parçalar: ZIP içindeki parçalar, numaralar atılarak kategorilere ayrılır
//    (word/header1.xml → word/header#.xml, xl/media/image3.png → medya). Aynı aileden
//    (OOXML→OOXML, ODF→ODF) bir kayıtta kaybolan kategori "kaybolan parça" sayılır.
//  - ogeler: biçimden bağımsız anlamlı öğe sayıları (paragraf, tablo, resim, yorum, dipnot,
//    slayt, birleşik hücre…). Biçimler arası (ör. .odt → .docx) karşılaştırmada da kullanılır.
//  - iliskiler: OOXML ilişki türü sayıları (.rels içindeki Type'ın son bölümü: styles, vbaProject,
//    customXml, digital-signature/origin…). Taşınmayan parçanın ilişkisi de kaybolur.
//  - ns: uzantı ad alanlarındaki öğe sayıları ("w14:checkbox", "x14:dataValidation",
//    "mc:AlternateContent", "dgm:dataModel"…). Ana ad alanları (w, x, p, a…) model düzeyinde ölçülür.
// Bağımlılık yoktur: ZIP merkez dizini elle okunur, sıkıştırma zlib ile açılır.
import { inflateRawSync } from "node:zlib";

const AZAMI_PARCA = 64 * 1024 * 1024; // daha büyük parça (ör. zip bombası) okunmaz

/**
 * ZIP merkez dizinindeki girdiler, sırasıyla ve yinelenenler dahil:
 * [{ad, yontem, sikisik, acik, yerel}]. ZIP değilse null.
 */
export function zipGirdileri(u8) {
  if (u8.length < 22 || u8[0] !== 0x50 || u8[1] !== 0x4b) return null;
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return null;
  const adet = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const dec = new TextDecoder();
  const girdiler = [];
  for (let k = 0; k < adet && p + 46 <= u8.length; k++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const adUz = dv.getUint16(p + 28, true);
    girdiler.push({
      ad: dec.decode(u8.subarray(p + 46, p + 46 + adUz)),
      yontem: dv.getUint16(p + 10, true),
      sikisik: dv.getUint32(p + 20, true),
      acik: dv.getUint32(p + 24, true),
      yerel: dv.getUint32(p + 42, true),
    });
    p += 46 + adUz + dv.getUint16(p + 30, true) + dv.getUint16(p + 32, true);
  }
  return girdiler;
}

/** ZIP merkez dizinini okur: Map<ad, () => Uint8Array|null>. ZIP değilse null. */
export function zipAc(u8) {
  const girdiler = zipGirdileri(u8);
  if (!girdiler) return null;
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  const parcalar = new Map();
  for (const { ad, yontem, sikisik, acik, yerel } of girdiler) {
    if (ad.endsWith("/")) continue;
    parcalar.set(ad, () => {
      if (acik > AZAMI_PARCA || yerel + 30 > u8.length) return null;
      const veriBas = yerel + 30 + dv.getUint16(yerel + 26, true) + dv.getUint16(yerel + 28, true);
      const ham = u8.subarray(veriBas, veriBas + sikisik);
      try {
        return yontem === 0 ? ham : yontem === 8 ? new Uint8Array(inflateRawSync(ham)) : null;
      } catch {
        return null;
      }
    });
  }
  return parcalar;
}

/** Parça adını kategoriye çevirir (numaralar #, medya tek kategori). */
export function kategori(ad) {
  if (/(^|\/)(media|Pictures)\//i.test(ad) || /\.(png|jpe?g|gif|bmp|tiff?|emf|wmf|svg)$/i.test(ad)) return "medya";
  if (/^Object \d+\//.test(ad) || /^Objects?\//.test(ad)) return "nesne/" + ad.replace(/^Object \d+\//, "").replace(/\d+/g, "#");
  return ad.replace(/\d+/g, "#");
}

/** OOXML [Content_Types].xml: {varsayilan: {uzantı: tür}, ozel: {/parça: tür}} ya da null. */
function icerikTurleri(xml) {
  if (!xml) return null;
  const varsayilan = {};
  const ozel = {};
  for (const m of xml.matchAll(/<Default [^>]*Extension="([^"]+)"[^>]*ContentType="([^"]+)"/g)) varsayilan[m[1].toLowerCase()] = m[2];
  for (const m of xml.matchAll(/<Override [^>]*PartName="([^"]+)"[^>]*ContentType="([^"]+)"/g)) ozel[m[1]] = m[2];
  return { varsayilan, ozel };
}

/**
 * OOXML parçasının kategorisi içerik türünden alınır (üreticiler aynı parçayı farklı yollara yazabilir:
 * xl/comments1.xml ↔ xl/comments/comment1.xml). İlişki (.rels) parçaları sayılmaz.
 */
function ooxmlKategori(ad, tur) {
  if (ad.endsWith(".rels") || ad === "[Content_Types].xml") return null;
  const ct = tur.ozel["/" + ad] ?? tur.varsayilan[ad.split(".").pop().toLowerCase()] ?? "";
  if (/^image\//.test(ct) || kategori(ad) === "medya") return "medya";
  if (!ct) return kategori(ad);
  // Ana belge parçasının türü değişkene göre değişir (xlsm/xltx → xlsx kaydı): tek kategori.
  if (/\.main\+xml$/.test(ct)) return "ana-belge";
  return ct
    .replace(/^application\/vnd\.openxmlformats-officedocument\./, "")
    .replace(/^application\/vnd\.openxmlformats-package\./, "")
    .replace(/^application\/vnd\.ms-(office|excel|word|powerpoint)\./, "ms-$1.")
    .replace(/^application\/vnd\.openxmlformats-/, "")
    .replace(/\+xml$/, "")
    .replace(/\.main$/, "");
}

/** Biçim ailesi: OOXML ya da ODF (paket karşılaştırması yalnız aynı aile içinde anlamlı). */
export function aile(parcalar) {
  if (!parcalar) return null;
  if (parcalar.has("[Content_Types].xml")) return "ooxml";
  if (parcalar.has("content.xml") || parcalar.has("mimetype")) return "odf";
  return "zip";
}

const say = (s, re) => (s.match(re) ?? []).length;

/** Anlamlı öğe sayıları; biçimden bağımsız anahtarlarla. */
function ogeSay(parcalar, metin) {
  const o = {};
  const ekle = (k, n) => {
    if (n) o[k] = (o[k] ?? 0) + n;
  };
  const adlar = [...parcalar.keys()];
  // Uyumluluk yedekleri (mc:Fallback içindeki VML/eski geçiş) aynı öğeyi iki kez saydırmasın.
  const yedeksiz = (x) => x.replace(/<mc:Fallback\b[\s\S]*?<\/mc:Fallback>/g, "");
  const metinA = metin;
  metin = (ad) => {
    const x = metinA(ad);
    return x && /\.xml$/.test(ad) && x.includes("mc:Fallback") ? yedeksiz(x) : x;
  };
  const xml = (re) => adlar.filter((a) => re.test(a)).map((a) => metin(a) ?? "");
  // ODF'de grafik nesnesinin yedek görüntüsü (ObjectReplacements) resim sayılmaz.
  const odfResim = (x) => say(x, /<draw:image (?![^>]*xlink:href="\.?\/?ObjectReplacements\/)/g);
  if (parcalar.has("word/document.xml")) {
    const d = metin("word/document.xml") ?? "";
    ekle("paragraf", say(d, /<w:p[ >\/]/g));
    ekle("tablo", say(d, /<w:tbl>/g));
    ekle("tabloSatırı", say(d, /<w:tr[ >]/g));
    ekle("köprü", say(d, /<w:hyperlink[ >]/g));
    ekle("değişiklik", say(d, /<w:(ins|del)[ >]/g));
    for (const x of xml(/^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/)) ekle("resim", say(x, /<pic:pic[ >]/g) + say(x, /<v:imagedata[ >]/g));
    ekle("dipnot", say(metin("word/footnotes.xml") ?? "", /<w:footnote (?![^>]*w:type=)/g));
    ekle("sonnot", say(metin("word/endnotes.xml") ?? "", /<w:endnote (?![^>]*w:type=)/g));
    ekle("yorum", say(metin("word/comments.xml") ?? "", /<w:comment[ >]/g));
    ekle("üstBilgi", adlar.filter((a) => /^word\/header\d*\.xml$/.test(a)).length);
    ekle("altBilgi", adlar.filter((a) => /^word\/footer\d*\.xml$/.test(a)).length);
  } else if (parcalar.has("xl/workbook.xml")) {
    const wb = metin("xl/workbook.xml") ?? "";
    ekle("sayfa", say(wb, /<sheet [^>]*\/?>/g));
    ekle("adlıAralık", say(wb, /<definedName[ >]/g));
    for (const s of xml(/^xl\/worksheets\/sheet\d+\.xml$/)) {
      ekle("hücre", say(s, /<c [^>]*>(?=\s*<[fvi])/g));
      ekle("formül", say(s, /<f[ >]/g));
      ekle("birleşik", say(s, /<mergeCell /g));
      ekle("köprü", say(s, /<hyperlink /g));
      ekle("doğrulama", say(s, /<dataValidation /g));
      ekle("koşullu", say(s, /<conditionalFormatting[ >]/g));
      ekle("filtre", say(s, /<autoFilter[ >]/g));
    }
    for (const c of xml(/^xl\/comments\d*\.xml$/)) ekle("yorum", say(c, /<comment /g));
    ekle("grafik", adlar.filter((a) => /^xl\/charts\/chart\d+\.xml$/.test(a)).length);
    for (const d of xml(/^xl\/drawings\/drawing\d+\.xml$/)) ekle("resim", say(d, /<xdr:pic>/g));
    ekle("tabloNesnesi", adlar.filter((a) => /^xl\/tables\/table\d+\.xml$/.test(a)).length);
    ekle("özetTablo", adlar.filter((a) => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(a)).length);
  } else if (parcalar.has("ppt/presentation.xml")) {
    ekle("slayt", say(metin("ppt/presentation.xml") ?? "", /<p:sldId /g));
    for (const s of xml(/^ppt\/slides\/slide\d+\.xml$/)) {
      ekle("şekil", say(s, /<p:sp>/g));
      ekle("resim", say(s, /<p:pic>/g));
      ekle("tablo", say(s, /<a:tbl>/g));
      ekle("grup", say(s, /<p:grpSp>/g));
      ekle("geçişliSlayt", /<p:transition[ >\/]/.test(s) ? 1 : 0);
    }
    ekle("grafik", adlar.filter((a) => /^ppt\/charts\/chart\d+\.xml$/.test(a)).length);
    ekle("not", adlar.filter((a) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(a)).length);
    for (const c of xml(/^ppt\/comments\/.*\.xml$/)) ekle("yorum", say(c, /<p:cm[ >]/g) + say(c, /<p188:cm[ >]/g));
  } else if (parcalar.has("content.xml")) {
    const c = metin("content.xml") ?? "";
    const st = metin("styles.xml") ?? "";
    if (/<office:text[ >]/.test(c)) {
      ekle("paragraf", say(c, /<text:(p|h)[ >\/]/g));
      ekle("tablo", say(c, /<table:table [^>]*>/g));
      ekle("tabloSatırı", say(c, /<table:table-row[ >\/]/g));
      ekle("köprü", say(c, /<text:a /g));
      ekle("değişiklik", say(c, /<text:changed-region /g));
      ekle("resim", odfResim(c));
      ekle("dipnot", say(c, /<text:note [^>]*text:note-class="footnote"/g));
      ekle("sonnot", say(c, /<text:note [^>]*text:note-class="endnote"/g));
      ekle("yorum", say(c, /<office:annotation[ >]/g));
      ekle("üstBilgi", say(st, /<style:header[ >]/g) ? 1 : 0);
      ekle("altBilgi", say(st, /<style:footer[ >]/g) ? 1 : 0);
    } else if (/<office:spreadsheet[ >]/.test(c)) {
      ekle("sayfa", say(c, /<table:table [^>]*>/g));
      ekle("adlıAralık", say(c, /<table:named-(range|expression) /g));
      for (const m of c.matchAll(/<table:table-cell ([^>]*)>/g)) {
        if (!/office:value-type=|table:formula=/.test(m[1])) continue;
        const tekrar = Number(/table:number-columns-repeated="(\d+)"/.exec(m[1])?.[1] ?? 1);
        ekle("hücre", Math.min(tekrar, 16384));
        if (/table:formula=/.test(m[1])) ekle("formül", Math.min(tekrar, 16384));
      }
      ekle("birleşik", say(c, /<table:table-cell [^>]*table:number-(?:columns|rows)-spanned="(?:[2-9]|\d{2,})"/g));
      ekle("köprü", say(c, /<text:a /g));
      ekle("doğrulama", say(c, /<table:content-validation /g));
      ekle("koşullu", say(c, /<calcext:conditional-format /g));
      ekle("yorum", say(c, /<office:annotation[ >]/g));
      ekle("grafik", say(c, /<draw:object /g));
      ekle("resim", odfResim(c));
      ekle("filtre", say(c, /table:display-filter-buttons="true"/g));
    } else if (/<office:presentation[ >]/.test(c)) {
      ekle("slayt", say(c, /<draw:page /g));
      ekle("şekil", say(c, /<draw:(custom-shape|rect|ellipse|line|connector|polygon|polyline|path) /g));
      ekle("resim", odfResim(c));
      ekle("tablo", say(c, /<table:table /g));
      ekle("grup", say(c, /<draw:g[ >]/g));
      ekle("grafik", say(c, /<draw:object /g));
      ekle("not", (c.match(/<presentation:notes[\s\S]*?<\/presentation:notes>/g) ?? []).filter((n) => /<text:p[^>]*>[^<]/.test(n) || /<text:span[^>]*>[^<]/.test(n)).length);
      ekle("yorum", say(c, /<officeooo:annotation[ >]/g) + say(c, /<office:annotation[ >]/g));
    }
  }
  return o;
}

/**
 * Uzantı ad alanları: URI → kısa önek. Envanter yalnız bunlardaki öğeleri sayar; ana ad alanlarının
 * (w, x, p, a, c, r, v…) öğeleri model ölçümüyle karşılaştırılır. Listede olmayan bir Microsoft uzantı
 * ad alanı (schemas.microsoft.com/office/…) yoldan türetilen önekle ("ms:powerpoint/2018/8/main") sayılır.
 */
export const UZANTI_NS = {
  "http://schemas.openxmlformats.org/markup-compatibility/2006": "mc",
  // Belge
  "http://schemas.microsoft.com/office/word/2006/wordml": "wne",
  "http://schemas.microsoft.com/office/word/2010/wordml": "w14",
  "http://schemas.microsoft.com/office/word/2012/wordml": "w15",
  "http://schemas.microsoft.com/office/word/2015/wordml/symex": "w16se",
  "http://schemas.microsoft.com/office/word/2016/wordml/cid": "w16cid",
  "http://schemas.microsoft.com/office/word/2018/wordml": "w16",
  "http://schemas.microsoft.com/office/word/2018/wordml/cex": "w16cex",
  "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash": "w16sdtdh",
  "http://schemas.microsoft.com/office/word/2023/wordml/word16du": "w16du",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing": "wp14",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingShape": "wps",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup": "wpg",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas": "wpc",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingInk": "wpi",
  // Tablo
  "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main": "x14",
  "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac": "x14ac",
  "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main": "x15",
  "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac": "x15ac",
  "http://schemas.microsoft.com/office/spreadsheetml/2014/11/main": "x16r2",
  "http://schemas.microsoft.com/office/spreadsheetml/2014/revision": "xr",
  "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2": "xr2",
  "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3": "xr3",
  "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata": "xlrd",
  "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2": "xlrd2",
  "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments": "tc",
  "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures": "xcalcf",
  "http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout": "xpdl",
  "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray": "xda",
  "http://schemas.microsoft.com/office/excel/2006/main": "xm",
  // Sunum
  "http://schemas.microsoft.com/office/powerpoint/2010/main": "p14",
  "http://schemas.microsoft.com/office/powerpoint/2012/main": "p15",
  "http://schemas.microsoft.com/office/powerpoint/2015/main": "p159",
  "http://schemas.microsoft.com/office/powerpoint/2017/10/main": "p1710",
  "http://schemas.microsoft.com/office/powerpoint/2018/8/main": "p188",
  "http://schemas.microsoft.com/office/powerpoint/2013/main/command": "pc",
  // Çizim, grafik, SmartArt
  "http://schemas.microsoft.com/office/drawing/2010/main": "a14",
  "http://schemas.microsoft.com/office/drawing/2012/main": "a15",
  "http://schemas.microsoft.com/office/drawing/2014/main": "a16",
  "http://schemas.microsoft.com/office/drawing/2016/SVG/main": "asvg",
  "http://schemas.microsoft.com/office/drawing/2017/decorative": "adec",
  "http://schemas.microsoft.com/office/drawing/2010/compatibility": "com14",
  "http://schemas.microsoft.com/office/drawing/2010/slicer": "sle15",
  "http://schemas.microsoft.com/office/drawing/2007/8/2/chart": "c14",
  "http://schemas.microsoft.com/office/drawing/2012/chart": "c15",
  "http://schemas.microsoft.com/office/drawing/2014/chart": "c16",
  "http://schemas.microsoft.com/office/drawing/2017/03/chart": "c16r3",
  "http://schemas.microsoft.com/office/drawing/2014/chartex": "cx",
  "http://schemas.microsoft.com/office/drawing/2012/chartStyle": "cs",
  "http://schemas.openxmlformats.org/drawingml/2006/diagram": "dgm",
  "http://purl.oclc.org/ooxml/drawingml/diagram": "dgm",
  "http://schemas.microsoft.com/office/drawing/2008/diagram": "dsp",
  "http://schemas.microsoft.com/office/drawing/2010/diagram": "dgm14",
  "http://schemas.microsoft.com/office/thememl/2012/main": "thm15",
  "http://schemas.microsoft.com/office/2006/activeX": "ax",
  "http://www.w3.org/2003/InkML": "inkml",
  // ODF (LibreOffice/OpenOffice uzantıları)
  "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0": "loext",
  "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0": "calcext",
  "http://openoffice.org/2009/office": "officeooo",
  "http://openoffice.org/2009/table": "tableooo",
  "http://openoffice.org/2010/draw": "drawooo",
  "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0": "field",
  "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0": "formx",
  "http://www.w3.org/TR/css3-text/": "css3t",
};

/** Ad alanı URI'sinin uzantı öneki; uzantı değilse null. */
export function uzantiOneki(uri) {
  if (UZANTI_NS[uri]) return UZANTI_NS[uri];
  const m = /^http:\/\/schemas\.microsoft\.com\/office\/(.+)$/.exec(uri);
  // İlişki ve (SharePoint) meta veri ad alanları uzantı öğesi taşımaz.
  if (m && !/relationships|metadata/.test(m[1])) return `ms:${m[1]}`;
  return null;
}

/** İlişki türünün kısa adı (http://…/relationships/styles → styles; …/metadata/core-properties → metadata/core-properties). */
export function iliskiTuru(tur) {
  const i = tur.lastIndexOf("/relationships/");
  return i >= 0 ? tur.slice(i + "/relationships/".length) : tur.slice(tur.lastIndexOf("/") + 1);
}

const XMLNS = /\bxmlns(?::([\w.-]+))?\s*=\s*["']([^"']*)["']/g;

/** Bir XML parçasındaki uzantı ad alanı öğelerini sayar (öneki bildirimden çözülür; kök varsayılan ad alanı da). */
function nsSay(x, ns) {
  const onekler = new Map();
  let varsayilan = null;
  for (const m of x.matchAll(XMLNS)) {
    const e = uzantiOneki(m[2]);
    if (m[1] === undefined) {
      if (varsayilan === null) varsayilan = e ?? "";
    } else if (e && !onekler.has(m[1])) onekler.set(m[1], e);
  }
  if (onekler.size) {
    const re = new RegExp(`<(${[...onekler.keys()].map((k) => k.replace(/\./g, "\\.")).join("|")}):([A-Za-z_][\\w.-]*)(?=[\\s/>])`, "g");
    for (const m of x.matchAll(re)) {
      const k = `${onekler.get(m[1])}:${m[2]}`;
      ns[k] = (ns[k] ?? 0) + 1;
    }
  }
  if (varsayilan) {
    for (const m of x.matchAll(/<([A-Za-z_][\w.-]*)(?=[\s/>])/g)) {
      const k = `${varsayilan}:${m[1]}`;
      ns[k] = (ns[k] ?? 0) + 1;
    }
  }
}

/** Paketteki ilişki türleri (OOXML) ve uzantı ad alanı öğeleri (OOXML ve ODF). */
function iliskiVeNs(p, metin) {
  const iliskiler = {};
  const ns = {};
  for (const ad of p.keys()) {
    if (ad.endsWith(".rels")) {
      for (const m of (metin(ad) ?? "").matchAll(/<Relationship\b[^>]*?\bType\s*=\s*["']([^"']+)["']/g)) {
        const k = iliskiTuru(m[1]);
        iliskiler[k] = (iliskiler[k] ?? 0) + 1;
      }
    } else if (/\.(xml|vml)$/i.test(ad) && ad !== "[Content_Types].xml") {
      const x = metin(ad);
      if (x) nsSay(x, ns);
    }
  }
  return { iliskiler, ns };
}

/** Bir dosyanın paket envanteri: {aile, parcalar, ogeler, iliskiler, ns} ({anahtar: adet}) ya da null (ZIP değilse). */
export function envanter(u8) {
  const p = zipAc(u8);
  if (!p) return null;
  const dec = new TextDecoder();
  const onbellek = new Map();
  const metin = (ad) => {
    if (!onbellek.has(ad)) {
      const b = p.get(ad)?.();
      onbellek.set(ad, b ? dec.decode(b) : null);
    }
    return onbellek.get(ad);
  };
  const parcalar = {};
  const tur = icerikTurleri(metin("[Content_Types].xml"));
  for (const ad of p.keys()) {
    const k = tur ? ooxmlKategori(ad, tur) : kategori(ad);
    if (k) parcalar[k] = (parcalar[k] ?? 0) + 1;
  }
  let ogeler = {};
  try {
    ogeler = ogeSay(p, metin);
  } catch {
    ogeler = {};
  }
  let ek = { iliskiler: {}, ns: {} };
  try {
    ek = iliskiVeNs(p, metin);
  } catch {
    /* bozuk parça: ilişki/ad alanı envanteri boş */
  }
  return { aile: aile(p), parcalar, ogeler, ...ek };
}

/** Envanter grupları ve kayıp/fark anahtarındaki önekleri. Parça, ilişki ve ad alanı yalnız aynı ailede karşılaştırılır. */
export const GRUPLAR = [
  ["parcalar", "parça", true],
  ["iliskiler", "ilişki", true],
  ["ns", "ns", true],
  ["ogeler", "öğe", false],
];

/**
 * İki envanter arasındaki kayıp: kaynakta olup hedefte azalan/kaybolan şeyler.
 * Parça kategorileri, ilişki türleri ve uzantı ad alanı öğeleri yalnız aynı ailede karşılaştırılır;
 * öğe sayıları her zaman.
 */
export function envanterKaybi(A, B) {
  const k = {};
  if (!A || !B) return k;
  for (const [grup, onEk, ayniAile] of GRUPLAR) {
    if (ayniAile && A.aile !== B.aile) continue;
    for (const [key, n] of Object.entries(A[grup] ?? {})) {
      const m = B[grup]?.[key] ?? 0;
      if (m < n) k[`${onEk}:${key}`] = n - m;
    }
  }
  return k;
}

/** İki envanter aynı mı (ikinci tur kararlılığı için). */
export function envanterFarki(B, C) {
  const f = {};
  if (!B || !C) return f;
  for (const [grup, onEk] of GRUPLAR) {
    const b = B[grup] ?? {};
    const c = C[grup] ?? {};
    for (const key of new Set([...Object.keys(b), ...Object.keys(c)])) {
      if ((b[key] ?? 0) !== (c[key] ?? 0)) f[`${onEk}:${key}`] = `${b[key] ?? 0}→${c[key] ?? 0}`;
    }
  }
  return f;
}
