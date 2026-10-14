// Temel çizgi: derlem turu, bulanık test ve LibreOffice denetiminin kaydedilmiş sonuçları ve
// yeni sonuçların bunlarla karşılaştırılması.
//
// Kurallar (kip, temel.json'daki "kip" alanından; --kip ya da DERLEM_KIP ile değiştirilebilir):
//   AĞIR (her kipte HATA): yeni takılma (süre aşımı), yığın taşması, bellek taşması, işçi çökmesi.
//   GERİLEME: yeni kayıp (bir kayıp sayacının artması), artık açılamayan / yeniden açılamayan dosya,
//             ikinci turda kararsızlaşma, formül sonucu farkı artışı, bulanık sınıfın kötüleşmesi
//             (ör. anlaşılır hata → ham hata), LibreOffice'in açamadığı yeni çıktı.
//             kip "baslangic": yalnız rapor edilir; kip "siki": HATA.
//   İYİLEŞME: azalan kayıp / düzelen dosya. CI başarısız olmaz; `pnpm derlem:temel` ile temel güncellenir.
//
// CLI:
//   node scripts/derlem/temel.mjs --guncelle [--zorla] [--kip siki|baslangic]
//       .derlem/rapor/*.json sonuçlarından temel.json'u yeniden yazar. Gerileme varsa --zorla ister.
//   node scripts/derlem/temel.mjs --calistir   (pnpm derlem:temel)
//       önce derlem turu + bulanık testi (LibreOffice varsa PDF denetimini de) çalıştırır, sonra günceller.
//   node scripts/derlem/temel.mjs --aday <dosya>
//       temel.json'a dokunmadan mevcut raporlardan aday temel yazar (CI bunu artefakt olarak yükler).
import { spawnSync } from "node:child_process";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { pathToFileURL } from "node:url";
import { argumanlar, KOK, RAPOR, TEMEL } from "./ortak.mjs";

export const AGIR_DURUM = new Set(["zaman-asimi", "bellek", "isci-hatasi", "cikis", "cokme"]);
export const BULANIK_SIRA = {
  açıldı: 0,
  uyarılı: 0,
  "anlaşılır hata": 0,
  "ham hata": 1,
  kaydedilemiyor: 1,
  "zaman aşımı": 2,
  takılma: 3,
  "yığın taşması": 3,
  "bellek taşması": 3,
  "işçi çöktü": 3,
};
export const BULANIK_AGIR = new Set(Object.entries(BULANIK_SIRA).filter(([, v]) => v === 3).map(([k]) => k));

export function temelOku() {
  if (!existsSync(TEMEL)) return { kip: "baslangic", derlem: {}, bulanik: { ornekler: {} }, lo: {} };
  return JSON.parse(readFileSync(TEMEL, "utf8"));
}
export const kipBelirle = (A, temel) => (typeof A.kip === "string" && A.kip) || process.env.DERLEM_KIP || temel.kip || "baslangic";

/* ------------------------------------------------------------ özetler (temel çizgide saklanan biçim) */

/** Tur sonuçlarından temel çizgi kaydı: dosya → {durum, açıldı, yığın, loFormülFarkı, turlar: {biçim: {hata, kayıp, kararsız}}}. */
export function derlemOzeti(dosyalar) {
  const o = {};
  for (const f of dosyalar) {
    const k = { durum: f.durum, acildi: !!f.acildi };
    if (f.yigin) k.yigin = true;
    if (f.loFormul) k.loFormulFarkli = f.loFormul.farkli;
    if (f.turlar?.length) {
      k.turlar = {};
      for (const t of f.turlar) {
        const r = {};
        if (t.hata) r.hata = true;
        if (t.yigin) r.yigin = true;
        if (Object.keys(t.kayip ?? {}).length) r.kayip = t.kayip;
        if (t.kararsiz?.length) r.kararsiz = true;
        k.turlar[t.bicim] = r;
      }
    }
    o[f.anahtar] = k;
  }
  return Object.fromEntries(Object.entries(o).sort(([a], [b]) => a.localeCompare(b)));
}

export function bulanikOzeti(sonuc) {
  const ornekler = {};
  for (const r of sonuc.ornekler) ornekler[`${r.app}/${r.ad}`] = r.sinif;
  return { hizMs: sonuc.hizMs, sinirMs: sonuc.temelSinirMs, ornekler: Object.fromEntries(Object.entries(ornekler).sort(([a], [b]) => a.localeCompare(b))) };
}

export function loOzeti(satirlar) {
  const o = {};
  for (const r of satirlar) o[r.anahtar] = { pdf: !r.ohHata, ...(r.bosFazla ? { bosFazla: r.bosFazla } : {}) };
  return Object.fromEntries(Object.entries(o).sort(([a], [b]) => a.localeCompare(b)));
}

/* ------------------------------------------------------------ karşılaştırma */

function bulgular(kip) {
  const b = { kip, hatalar: [], gerilemeler: [], iyilesmeler: [], bilgiler: [] };
  b.gerileme = (s) => (kip === "siki" ? b.hatalar : b.gerilemeler).push(s);
  return b;
}
const agirMi = (x) => x && (AGIR_DURUM.has(x.durum) || x.yigin || Object.values(x.turlar ?? {}).some((t) => t.yigin));
const agirAciklama = (x) => (x.yigin || Object.values(x.turlar ?? {}).some((t) => t.yigin) ? "yığın taşması" : { "zaman-asimi": "takılma (süre aşımı)", bellek: "bellek taşması", "isci-hatasi": "işçi hatası", cikis: "işçi çıktı", cokme: "çökme" }[x.durum] ?? x.durum);

export function derlemKarsilastir(simdi, temel, kip) {
  const b = bulgular(kip);
  for (const [k, s] of Object.entries(simdi)) {
    const t = temel[k];
    if (agirMi(s) && !(t && agirMi(t) && agirAciklama(t) === agirAciklama(s))) b.hatalar.push(`${k}: ${agirAciklama(s)}`);
    if (!t) {
      b.bilgiler.push(`${k}: temelde yok (yeni dosya; ${s.acildi ? "açıldı" : "açılamadı"})`);
      continue;
    }
    if (t.acildi && !s.acildi && !agirMi(s)) b.gerileme(`${k}: artık açılamıyor`);
    if (!t.acildi && s.acildi) b.iyilesmeler.push(`${k}: artık açılıyor`);
    if (agirMi(t) && !agirMi(s)) b.iyilesmeler.push(`${k}: ${agirAciklama(t)} giderildi`);
    if ((s.loFormulFarkli ?? 0) > (t.loFormulFarkli ?? 0)) b.gerileme(`${k}: LibreOffice'ten farklı formül sonucu ${t.loFormulFarkli ?? 0}→${s.loFormulFarkli}`);
    if ((s.loFormulFarkli ?? 0) < (t.loFormulFarkli ?? 0)) b.iyilesmeler.push(`${k}: formül sonucu farkı ${t.loFormulFarkli}→${s.loFormulFarkli ?? 0}`);
    for (const [bicim, sr] of Object.entries(s.turlar ?? {})) {
      const tr = t.turlar?.[bicim];
      if (!tr) continue;
      if (sr.hata && !tr.hata) b.gerileme(`${k} →${bicim}: kaydedilemiyor ya da yeniden açılamıyor (yeni)`);
      if (!sr.hata && tr.hata) b.iyilesmeler.push(`${k} →${bicim}: artık kaydedilip yeniden açılıyor`);
      if (sr.kararsiz && !tr.kararsiz) b.gerileme(`${k} →${bicim}: ikinci kayıtta değişiyor (yeni)`);
      if (!sr.kararsiz && tr.kararsiz) b.iyilesmeler.push(`${k} →${bicim}: ikinci kayıtta kararlı`);
      if (sr.hata || tr.hata) continue;
      const artan = [];
      const azalan = [];
      for (const key of new Set([...Object.keys(sr.kayip ?? {}), ...Object.keys(tr.kayip ?? {})])) {
        const a = tr.kayip?.[key] ?? 0;
        const n = sr.kayip?.[key] ?? 0;
        if (n > a) artan.push(`${key} ${a}→${n}`);
        if (n < a) azalan.push(`${key} ${a}→${n}`);
      }
      if (artan.length) b.gerileme(`${k} →${bicim}: yeni kayıp: ${artan.join(", ")}`);
      if (azalan.length) b.iyilesmeler.push(`${k} →${bicim}: kayıp azaldı: ${azalan.join(", ")}`);
    }
  }
  for (const k of Object.keys(temel)) if (!simdi[k]) b.bilgiler.push(`${k}: bu koşuda yok (atlandı)`);
  return b;
}

export function bulanikKarsilastir(simdi, temel, kip) {
  const b = bulgular(kip);
  const T = temel?.ornekler ?? {};
  for (const [k, sinif] of Object.entries(simdi.ornekler)) {
    const t = T[k];
    if (BULANIK_AGIR.has(sinif) && t !== sinif) b.hatalar.push(`${k}: ${sinif}${t ? ` (önce: ${t})` : " (yeni örnek)"}`);
    if (!t) continue;
    const ss = BULANIK_SIRA[sinif] ?? 1;
    const ts = BULANIK_SIRA[t] ?? 1;
    if (ss > ts && !BULANIK_AGIR.has(sinif)) b.gerileme(`${k}: ${t} → ${sinif}`);
    else if (ss < ts) b.iyilesmeler.push(`${k}: ${t} → ${sinif}`);
    else if (sinif !== t) b.bilgiler.push(`${k}: ${t} → ${sinif}`);
  }
  for (const k of Object.keys(T)) if (!simdi.ornekler[k]) b.bilgiler.push(`${k}: bu koşuda yok`);
  return b;
}

export function loKarsilastir(simdi, temel, kip) {
  const b = bulgular(kip);
  for (const [k, s] of Object.entries(simdi)) {
    const t = temel?.[k];
    if (!t) continue;
    if (t.pdf && !s.pdf) b.gerileme(`${k}: LibreOffice artık açamıyor (PDF üretilemedi)`);
    if (!t.pdf && s.pdf) b.iyilesmeler.push(`${k}: LibreOffice artık açıyor`);
    if ((s.bosFazla ?? 0) > (t.bosFazla ?? 0)) b.gerileme(`${k}: fazladan boş sayfa ${t.bosFazla ?? 0}→${s.bosFazla}`);
    if ((s.bosFazla ?? 0) < (t.bosFazla ?? 0)) b.iyilesmeler.push(`${k}: fazladan boş sayfa ${t.bosFazla}→${s.bosFazla ?? 0}`);
  }
  return b;
}

/** Bulguları Markdown'a çevirir ve konsola özetler; dönüş: hata var mı. */
export function bulguRaporu(baslik, b, { azami = 60 } = {}) {
  const L = [`## ${baslik} — temel çizgiyle karşılaştırma (kip: ${b.kip})`, ""];
  const liste = (ad, a) => {
    L.push(`### ${ad} (${a.length})`, "");
    for (const s of a.slice(0, azami)) L.push(`- ${s}`);
    if (a.length > azami) L.push(`- … ve ${a.length - azami} tane daha`);
    L.push("");
  };
  liste("HATA", b.hatalar);
  liste(b.kip === "siki" ? "Gerileme (HATA sayıldı)" : "Gerileme (yalnız rapor)", b.gerilemeler);
  liste("İyileşme", b.iyilesmeler);
  if (b.bilgiler.length) liste("Bilgi", b.bilgiler);
  if (b.iyilesmeler.length) L.push("> İyileşme var: `pnpm derlem:temel` ile temel çizgiyi güncelleyin (scripts/derlem/temel.json).", "");
  const md = L.join("\n");
  console.log(`\n${baslik}: ${b.hatalar.length} hata, ${b.gerilemeler.length} gerileme (rapor), ${b.iyilesmeler.length} iyileşme, ${b.bilgiler.length} bilgi (kip: ${b.kip})`);
  for (const s of b.hatalar.slice(0, 40)) console.log(`  HATA: ${s}`);
  for (const s of b.gerilemeler.slice(0, 15)) console.log(`  gerileme: ${s}`);
  if (b.iyilesmeler.length) console.log(`  ${b.iyilesmeler.length} iyileşme → pnpm derlem:temel`);
  return { md, hataVar: b.hatalar.length > 0 };
}

/* ------------------------------------------------------------ CLI */
async function ana() {
  const A = argumanlar();
  if (A.calistir) {
    const kos = (betik, ek = []) => spawnSync(process.execPath, [join(KOK, betik), "--temelsiz", ...ek], { stdio: "inherit", cwd: KOK }).status;
    kos("scripts/derlem/tur.mjs");
    kos("scripts/bulanik/calistir.mjs", ["--paketleme-yok"]);
    if (spawnSync("soffice", ["--version"], { stdio: "ignore" }).status === 0) kos("scripts/derlem/lo.mjs");
  }
  if (!A.guncelle && !A.calistir && !A.aday) {
    console.log("Kullanım: node scripts/derlem/temel.mjs --guncelle [--zorla] [--kip siki|baslangic]  |  --calistir  |  --aday <dosya>");
    return 0;
  }
  const eski = temelOku();
  const yeni = { ...eski };
  const oku = (ad) => (existsSync(join(RAPOR, ad)) ? JSON.parse(readFileSync(join(RAPOR, ad), "utf8")) : null);
  const tur = oku("derlem.json");
  const bul = oku("bulanik.json");
  // LibreOffice denetimi bu koşuda yapılmadıysa eski raporu kullanma (eski temel korunur).
  const lo = A.calistir && spawnSync("soffice", ["--version"], { stdio: "ignore" }).status !== 0 ? null : oku("lo_pdf.json");
  for (const [ad, r] of [["derlem", tur], ["bulanık", bul]]) {
    if (r && r.kapsamli === false) {
      console.error(`Temel güncellenmedi: ${ad} raporu kısmi (--uygulama/--dosya ile koşulmuş). Tam koşu gerekir.`);
      return 1;
    }
  }
  if (!tur && !bul) {
    console.error("Rapor yok: önce `pnpm derlem` ve `pnpm bulanik` çalıştırın (ya da `pnpm derlem:temel`).");
    return 1;
  }
  const kip = A.kip ?? eski.kip ?? "baslangic";
  const gerilemeler = [];
  if (tur) {
    const oz = derlemOzeti(tur.dosyalar);
    const b = derlemKarsilastir(oz, eski.derlem ?? {}, "siki");
    gerilemeler.push(...b.hatalar);
    // Bu makinede üretilemeyen (LibreOffice yok) büyük dosyaların eski kayıtları korunur.
    const korunan = Object.entries(eski.derlem ?? {}).filter(([k]) => k.startsWith("uretilen/") && !oz[k]);
    yeni.derlem = Object.fromEntries([...Object.entries(oz), ...korunan].sort(([a], [b]) => a.localeCompare(b)));
    yeni.derlemBilgi = { tarih: tur.tarih, commit: tur.commit, dosya: tur.dosyalar.length };
  }
  if (bul) {
    const oz = bulanikOzeti(bul);
    const b = bulanikKarsilastir(oz, eski.bulanik, "siki");
    gerilemeler.push(...b.hatalar);
    yeni.bulanik = oz;
  }
  if (lo) {
    const oz = loOzeti(lo.satirlar);
    const b = loKarsilastir(oz, eski.lo, "siki");
    gerilemeler.push(...b.hatalar);
    yeni.lo = oz;
  }
  if (gerilemeler.length && !A.zorla && !A.aday && Object.keys(eski.derlem ?? {}).length) {
    console.error(`Temel güncellenmedi: ${gerilemeler.length} gerileme var (yalnız iyileşmeler kendiliğinden kabul edilir). İlk 20:`);
    for (const s of gerilemeler.slice(0, 20)) console.error(`  - ${s}`);
    console.error("Bilerek kabul etmek için: node scripts/derlem/temel.mjs --guncelle --zorla");
    return 1;
  }
  const sirali = {
    aciklama: "Derlem turu / bulanık test / LibreOffice denetimi temel çizgisi. Elle düzenlemeyin: pnpm derlem:temel ile güncellenir. Kurallar: scripts/derlem/temel.mjs başı.",
    kip,
    guncelleme: new Date().toISOString().slice(0, 10),
    derlemBilgi: yeni.derlemBilgi,
    bulanik: yeni.bulanik,
    lo: yeni.lo ?? {},
    derlem: yeni.derlem ?? {},
  };
  // --aday: temel.json'a dokunmadan aday temel yazılır (CI artefaktı; koordinatör inceleyip kopyalar).
  const hedef = typeof A.aday === "string" ? A.aday : TEMEL;
  writeFileSync(hedef, JSON.stringify(sirali, null, 1) + "\n");
  if (A.aday) console.log(`aday temel yazıldı: ${hedef} (${gerilemeler.length} gerileme içerir)`);
  else console.log(`temel çizgi yazıldı: ${TEMEL} (kip: ${kip}; derlem ${Object.keys(sirali.derlem).length} dosya, bulanık ${Object.keys(sirali.bulanik?.ornekler ?? {}).length} örnek, LO ${Object.keys(sirali.lo).length} çıktı)`);
  return 0;
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) process.exitCode = await ana();
