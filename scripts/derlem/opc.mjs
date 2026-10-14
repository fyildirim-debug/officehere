// Paket tutarlılık denetimi (pnpm derlem:opc).
//
// OOXML (OPC) paketlerinde:
//   - [Content_Types].xml: var mı, her parçanın içerik türü var mı, var olmayan parçaya Override, yinelenen
//     Default/Override (ZIP'te ilk girdi olması OPC'nin şartı değil; LibreOffice de sona yazar, denetlenmez)
//   - parçalar: yinelenen ZIP girdisi (OPC'de parça adları büyük/küçük harf duyarsızdır), geçersiz parça adı
//   - ilişkiler: eksik öznitelik, yinelenen Id, hedefi olmayan iç ilişki, kaynağı olmayan .rels, XML'de
//     .rels'te karşılığı olmayan ilişki kimliği (r:id, r:embed, r:dm…, VML o:relid: kırık rId)
//   - ana belge ilişkisi ve dosya uzantısıyla ana parçanın içerik türünün uyumu (.docx ↔ makrolu tür)
// ODF paketlerinde: mimetype (ilk girdi, sıkıştırmasız, manifest kök türüyle aynı) ve manifest ↔ ZIP girdileri.
//
// Kullanım:
//   pnpm derlem:opc [--temelsiz] [--kip …]
//       derlem turunun kaydettiği çıktılar (.derlem/rapor/derlem.json) ve özgünleri denetlenir; çıktıların
//       hata sayıları temel çizgiyle (temel.json "opc") karşılaştırılır: artış her kipte HATA, azalma iyileşme.
//       Rapor: .derlem/rapor/opc.json, opc.md
//   node scripts/derlem/opc.mjs <dosya|dizin>...   yalnız listeler (temel karşılaştırması yok)
import { existsSync, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { extname, join, relative } from "node:path";
import { pathToFileURL } from "node:url";
import { zipAc, zipGirdileri } from "./envanter.mjs";
import { argumanlar, KOK, posix, RAPOR } from "./ortak.mjs";
import { bulguRaporu, hataSayilari, kipBelirle, opcKarsilastir, opcOzeti, temelOku } from "./temel.mjs";

const ILISKI_NS = new Set(["http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships"]);
const VML_OFIS_NS = "urn:schemas-microsoft-com:office:office";

/** Uzantı → ana parçanın içerik türü. */
const ANA_TUR = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  docm: "application/vnd.ms-word.document.macroEnabled.main+xml",
  dotx: "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
  dotm: "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  xlsm: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
  xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
  xltm: "application/vnd.ms-excel.template.macroEnabled.main+xml",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
  pptm: "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
  potx: "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
  potm: "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
  ppsx: "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
  ppsm: "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
};
/** ODF uzantısı → mimetype. */
const ODF_TUR = {
  odt: "application/vnd.oasis.opendocument.text",
  ott: "application/vnd.oasis.opendocument.text-template",
  ods: "application/vnd.oasis.opendocument.spreadsheet",
  ots: "application/vnd.oasis.opendocument.spreadsheet-template",
  odp: "application/vnd.oasis.opendocument.presentation",
  otp: "application/vnd.oasis.opendocument.presentation-template",
};
export const DENETLENEN_UZANTILAR = new Set([...Object.keys(ANA_TUR), ...Object.keys(ODF_TUR)]);

const kucuk = (s) => s.toLowerCase(); // OPC parça adı karşılaştırması ASCII büyük/küçük harf duyarsızdır
const oznitelik = (s, ad) => new RegExp(`(?:^|\\s)${ad}\\s*=\\s*("([^"]*)"|'([^']*)')`).exec(s)?.slice(2).find((x) => x !== undefined);
const xmlCoz = (s) => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
function yuzdeCoz(s) {
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}

/** Göreli ilişki hedefini kaynak parçanın dizinine göre paket içi ada çevirir (fragman atılır). */
export function hedefCoz(kaynakParca, hedef) {
  const temiz = yuzdeCoz(xmlCoz(hedef).split("#")[0]);
  if (temiz.startsWith("/")) return temiz.slice(1);
  const yol = kaynakParca.split("/").slice(0, -1);
  for (const s of temiz.split("/")) {
    if (s === "..") yol.pop();
    else if (s !== "." && s !== "") yol.push(s);
  }
  return yol.join("/");
}

/**
 * Bir paketin tutarlılık hataları: [{tur, ayrinti}]. `uzanti` verilirse ana parçanın içerik türü
 * (OOXML) ya da mimetype (ODF) uzantıyla karşılaştırılır. ZIP olmayan girdi için tek hata döner.
 */
export function opcDenetle(u8, { uzanti } = {}) {
  const h = [];
  const ekle = (tur, ayrinti) => h.push({ tur, ayrinti });
  const girdiler = zipGirdileri(u8);
  const z = girdiler && zipAc(u8);
  if (!z) return [{ tur: "zip-okunamadı", ayrinti: "ZIP merkez dizini bulunamadı" }];
  const dec = new TextDecoder();
  const metin = (ad) => {
    const b = z.get(ad)?.();
    return b ? dec.decode(b) : null;
  };
  const dosyalar = girdiler.filter((g) => !g.ad.endsWith("/"));
  const uz = uzanti ? kucuk(uzanti).replace(/^\./, "") : null;

  // Yinelenen girdi: aynı ad ya da yalnız harf büyüklüğü farklı ad
  const gorulen = new Map();
  for (const g of dosyalar) {
    const k = kucuk(g.ad);
    if (gorulen.has(k)) ekle("yinelenen-parça", gorulen.get(k) === g.ad ? g.ad : `${gorulen.get(k)} ↔ ${g.ad}`);
    else gorulen.set(k, g.ad);
  }

  if (z.has("mimetype") || (z.has("META-INF/manifest.xml") && !z.has("[Content_Types].xml"))) {
    odfDenetle(girdiler, z, metin, uz, ekle);
    return h;
  }

  const adlar = new Set(dosyalar.map((g) => kucuk(g.ad)));
  for (const g of dosyalar) {
    if (g.ad === "[Content_Types].xml") continue;
    if (g.ad.startsWith("/") || g.ad.includes("\\") || g.ad.split("/").some((s) => s === "" || s.endsWith("."))) ekle("geçersiz-parça-adı", g.ad);
  }

  // [Content_Types].xml
  const ct = metin("[Content_Types].xml");
  const varsayilan = new Map();
  const ozel = new Map();
  if (ct === null) ekle("içerik-türleri-yok", "[Content_Types].xml");
  else {
    for (const m of ct.matchAll(/<Default\b([^>]*)>/g)) {
      const e = oznitelik(m[1], "Extension");
      const t = oznitelik(m[1], "ContentType");
      if (!e || !t) {
        ekle("içerik-türü-özniteliği-eksik", m[0].slice(0, 80));
        continue;
      }
      if (varsayilan.has(kucuk(e))) ekle("yinelenen-içerik-türü", `Default ${e}`);
      varsayilan.set(kucuk(e), t);
    }
    for (const m of ct.matchAll(/<Override\b([^>]*)>/g)) {
      const p = oznitelik(m[1], "PartName");
      const t = oznitelik(m[1], "ContentType");
      if (!p || !t) {
        ekle("içerik-türü-özniteliği-eksik", m[0].slice(0, 80));
        continue;
      }
      const ad = yuzdeCoz(p.replace(/^\//, ""));
      if (ozel.has(kucuk(ad))) ekle("yinelenen-içerik-türü", `Override ${p}`);
      ozel.set(kucuk(ad), t);
      if (!adlar.has(kucuk(ad))) ekle("fazla-override", p);
    }
    for (const g of dosyalar) {
      if (g.ad === "[Content_Types].xml") continue;
      const e = g.ad.includes(".") ? kucuk(g.ad.slice(g.ad.lastIndexOf(".") + 1)) : "";
      if (!ozel.has(kucuk(g.ad)) && !varsayilan.has(e)) ekle("içerik-türü-yok", g.ad);
    }
  }
  const icerikTuru = (ad) => ozel.get(kucuk(ad)) ?? varsayilan.get(ad.includes(".") ? kucuk(ad.slice(ad.lastIndexOf(".") + 1)) : "");

  // İlişkiler
  const kimlikler = new Map(); // kaynak parça (küçük harf) → Set<Id>
  let anaParca = null;
  for (const g of dosyalar) {
    if (!/\.rels$/i.test(g.ad)) continue;
    const m = /^(.*?)_rels\/([^/]*)\.rels$/i.exec(g.ad);
    if (!m) {
      ekle("geçersiz-rels-konumu", g.ad);
      continue;
    }
    const kaynak = m[1] + m[2];
    const kok = g.ad === "_rels/.rels";
    if (!kok && !adlar.has(kucuk(kaynak))) ekle("sahipsiz-rels", g.ad);
    const x = metin(g.ad) ?? "";
    const ids = new Set();
    kimlikler.set(kucuk(kaynak), ids);
    for (const r of x.matchAll(/<Relationship\b([^>]*?)\/?>/g)) {
      const id = oznitelik(r[1], "Id");
      const tur = oznitelik(r[1], "Type");
      const hedef = oznitelik(r[1], "Target");
      if (!id || !tur || hedef === undefined) {
        ekle("ilişki-özniteliği-eksik", `${g.ad}: ${r[0].slice(0, 100)}`);
        continue;
      }
      if (ids.has(id)) ekle("yinelenen-rId", `${g.ad}: ${id}`);
      ids.add(id);
      // Dış hedef ve yalnız fragmandan oluşan iç bağlantı (#Sayfa1!A1, #_ftn1) parça göstermez.
      if (/^external$/i.test(oznitelik(r[1], "TargetMode") ?? "") || hedef.startsWith("#")) continue;
      const yol = hedefCoz(kok ? "" : kaynak, hedef);
      if (!yol || !adlar.has(kucuk(yol))) ekle("hedef-yok", `${g.ad} → ${hedef}`);
      else if (kok && /\/officeDocument$/.test(tur)) anaParca ??= yol;
    }
  }
  if (ct !== null) {
    if (!anaParca) ekle("ana-belge-yok", "_rels/.rels içinde officeDocument ilişkisi yok");
    else if (uz && ANA_TUR[uz] && kucuk(icerikTuru(anaParca) ?? "") !== kucuk(ANA_TUR[uz])) ekle("uzantı-içerik-türü", `.${uz}: ${anaParca} = ${icerikTuru(anaParca) ?? "-"}`);
  }

  // Kırık rId: XML'deki ilişki kimliği parçanın .rels'inde yok
  for (const g of dosyalar) {
    if (!/\.(xml|vml)$/i.test(g.ad) || g.ad === "[Content_Types].xml") continue;
    const x = metin(g.ad);
    if (!x) continue;
    const onekler = [];
    for (const m of x.matchAll(/\bxmlns:([\w.-]+)\s*=\s*["']([^"']*)["']/g)) if (ILISKI_NS.has(m[2]) || m[2] === VML_OFIS_NS) onekler.push([m[1], m[2] === VML_OFIS_NS]);
    if (!onekler.length) continue;
    const ids = kimlikler.get(kucuk(g.ad));
    const eksik = new Set();
    for (const [onek, vml] of onekler) {
      const re = vml ? new RegExp(`\\s${onek.replace(/\./g, "\\.")}:relid\\s*=\\s*["']([^"']*)["']`, "g") : new RegExp(`\\s${onek.replace(/\./g, "\\.")}:[A-Za-z]+\\s*=\\s*["']([^"']*)["']`, "g");
      for (const m of x.matchAll(re)) if (m[1] && !ids?.has(m[1])) eksik.add(m[1]);
    }
    for (const id of eksik) ekle("kırık-rId", `${g.ad}: ${id}`);
  }
  return h;
}

function odfDenetle(girdiler, z, metin, uz, ekle) {
  const ilk = girdiler.find((g) => !g.ad.endsWith("/"));
  const mt = metin("mimetype");
  if (mt === null) ekle("odf-mimetype", "mimetype yok");
  else {
    if (ilk?.ad !== "mimetype") ekle("odf-mimetype", `ilk girdi: ${ilk?.ad}`);
    else if (ilk.yontem !== 0) ekle("odf-mimetype", "mimetype sıkıştırılmış");
    if (uz && ODF_TUR[uz] && mt.trim() !== ODF_TUR[uz]) ekle("uzantı-içerik-türü", `.${uz}: ${mt.trim()}`);
  }
  const man = metin("META-INF/manifest.xml");
  if (man === null) {
    ekle("odf-manifest-yok", "META-INF/manifest.xml");
    return;
  }
  const listede = new Set();
  for (const m of man.matchAll(/<manifest:file-entry\b([^>]*?)\/?>/g)) {
    const yol = oznitelik(m[1], "manifest:full-path");
    if (yol === undefined) continue;
    const ad = xmlCoz(yol);
    listede.add(ad);
    if (ad === "/") {
      const tur = oznitelik(m[1], "manifest:media-type");
      if (mt !== null && tur !== undefined && xmlCoz(tur) !== mt.trim()) ekle("odf-mimetype", `manifest kök türü ${tur} ≠ ${mt.trim()}`);
    } else if (!ad.endsWith("/") && !z.has(ad)) ekle("odf-manifest-fazla", ad);
  }
  for (const g of girdiler) {
    if (g.ad.endsWith("/") || g.ad === "mimetype" || g.ad.startsWith("META-INF/")) continue;
    if (!listede.has(g.ad)) ekle("odf-manifest-eksik", g.ad);
  }
}

/* ------------------------------------------------------------ CLI */

function dosyaListesi(yollar) {
  const out = [];
  for (const y of yollar) {
    if (!existsSync(y)) continue;
    if (statSync(y).isDirectory()) out.push(...dosyaListesi(readdirSync(y).sort().map((f) => join(y, f))));
    else if (DENETLENEN_UZANTILAR.has(extname(y).slice(1).toLowerCase())) out.push(y);
  }
  return out;
}
const denetle = (yol) => opcDenetle(new Uint8Array(readFileSync(yol)), { uzanti: extname(yol) });

async function ana() {
  const A = argumanlar();
  if (A._.length) {
    let toplam = 0;
    for (const f of dosyaListesi(A._)) {
      const h = denetle(f);
      toplam += h.length;
      console.log(`${posix(relative(process.cwd(), f))}: ${h.length ? h.map((x) => `${x.tur} (${x.ayrinti})`).slice(0, 8).join(" | ") + (h.length > 8 ? ` (+${h.length - 8})` : "") : "tamam"}`);
    }
    console.log(`\n${toplam} paket hatası`);
    return 0;
  }
  const raporYolu = join(RAPOR, "derlem.json");
  if (!existsSync(raporYolu)) {
    console.error("Önce `pnpm derlem` çalıştırın (kaydedilen dosyalar ve derlem.json gerekli).");
    return 1;
  }
  const t0 = performance.now();
  const tur = JSON.parse(readFileSync(raporYolu, "utf8"));
  const satirlar = [];
  for (const f of tur.dosyalar) {
    const kaynakUz = extname(f.yol).slice(1).toLowerCase();
    const ozgun = DENETLENEN_UZANTILAR.has(kaynakUz) && existsSync(f.yol) ? denetle(f.yol) : null;
    for (const t of f.turlar ?? []) {
      if (t.hata || !t.ciktiYolu || !DENETLENEN_UZANTILAR.has(t.bicim) || !existsSync(t.ciktiYolu)) continue;
      satirlar.push({ anahtar: `${f.anahtar} →${t.bicim}`, app: f.app, cikti: posix(relative(KOK, t.ciktiYolu)), hatalar: denetle(t.ciktiYolu), ozgun: ozgun ? hataSayilari(ozgun) : null });
    }
  }
  const sureMs = Math.round(performance.now() - t0);
  writeFileSync(join(RAPOR, "opc.json"), JSON.stringify({ tarih: new Date().toISOString(), derlemTarih: tur.tarih, commit: tur.commit, kapsamli: tur.kapsamli, sureMs, satirlar }, null, 1));

  const L = ["# Paket tutarlılık denetimi (OPC / ODF)", "", `${satirlar.length} çıktı · ${sureMs} ms`, ""];
  L.push("| uygulama | çıktı | hatalı çıktı | hata | özgünü hatalı |", "|---|---|---|---|---|");
  for (const app of [...new Set(satirlar.map((s) => s.app))]) {
    const S = satirlar.filter((s) => s.app === app);
    L.push(`| ${app} | ${S.length} | ${S.filter((s) => s.hatalar.length).length} | ${S.reduce((a, s) => a + s.hatalar.length, 0)} | ${S.filter((s) => s.ozgun && Object.keys(s.ozgun).length).length} |`);
  }
  const turToplam = hataSayilari(satirlar.flatMap((s) => s.hatalar));
  L.push("", "Hata türleri: " + (Object.entries(turToplam).map(([k, v]) => `\`${k}\` ${v}`).join(", ") || "yok"), "");
  const hatali = satirlar.filter((s) => s.hatalar.length);
  if (hatali.length) {
    L.push("| çıktı | hatalar | örnek | özgünde |", "|---|---|---|---|");
    for (const s of hatali) {
      const ornek = s.hatalar.slice(0, 3).map((x) => x.ayrinti).join("; ").replace(/\|/g, "/").slice(0, 160);
      L.push(`| ${s.anahtar} | ${Object.entries(hataSayilari(s.hatalar)).map(([k, v]) => `${k}:${v}`).join(" ")} | ${ornek} | ${s.ozgun ? Object.entries(s.ozgun).map(([k, v]) => `${k}:${v}`).join(" ") || "—" : "-"} |`);
    }
  }
  let md = L.join("\n") + "\n";
  let kod = 0;
  if (!A.temelsiz) {
    const temel = temelOku();
    const b = opcKarsilastir(opcOzeti(satirlar), temel.opc ?? {}, satirlar.map((s) => s.anahtar), kipBelirle(A, temel));
    const { md: kmd, hataVar } = bulguRaporu("Paket tutarlılık denetimi", b);
    md += "\n" + kmd + "\n";
    if (hataVar) kod = 1;
  }
  writeFileSync(join(RAPOR, "opc.md"), md);
  console.log(`\nPaket denetimi: ${satirlar.length} çıktı, ${hatali.length} hatalı, ${satirlar.reduce((a, s) => a + s.hatalar.length, 0)} hata (${sureMs} ms). Rapor: ${join(RAPOR, "opc.md")}`);
  return kod;
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) process.exitCode = await ana();
