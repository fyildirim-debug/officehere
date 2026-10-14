// Kaydet-aç turu (pnpm derlem).
//
// Derlemdeki her dosya uygulamanın kendi okuyucusuyla açılır, yerel biçimde (ve yazılabiliyorsa kendi
// biçiminde) kaydedilir, yeniden açılır ve karşılaştırılır:
//   - model: metin/kelime, düğüm ve işaretler, hücre değeri/formül/etkin stil, sayfa yapısı,
//            slayt/şekil/not/geçiş, belge ayarları (üst/alt bilgi, yorum, sayfa)
//   - paket: OOXML/ODF parça envanteri ve anlamlı öğe sayıları (envanter.mjs)
//   - ikinci tur: kaydedilen dosya bir kez daha yazılıp açılınca değişiyor mu (yakınsama)
//   - tablo: motorun formül sonuçları LibreOffice'in hesabıyla (test-derlem/gercek_lo.json)
//   - POI derlemi (test-derlem/poi/, gerçek Office dosyaları): uygulama başına parça/ilişki/ad alanı kaybı
//     toplamı; temel.json "poi" ile karşılaştırılır, parça kaybı artışı her kipte HATA (temel.mjs)
// Çıktılar: .derlem/rapor/derlem.json (ham), derlem.md (özet + "kaybolan" sayıları) ve temel çizgi
// karşılaştırması (temel.mjs); kaydedilen dosyalar .derlem/cikti/<uygulama>/ (pnpm derlem:lo bunları açar).
//
// Kullanım: pnpm derlem [--uygulama belge|tablo|sunum] [--dosya <alt dize>] [--temelsiz] [--kip siki|baslangic]
import { execSync } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { availableParallelism } from "node:os";
import { basename, dirname, join } from "node:path";
import { FARK, gercekKarsilastir, kararsizlik, kayipHaritasi } from "./fark.mjs";
import { Havuz } from "./havuz.mjs";
import { argumanlar, CIKTI, DERLEM, derlemDosyalari, KOK, PAKET, RAPOR, sure, URETILEN, UYGULAMALAR } from "./ortak.mjs";
import { paketle } from "./paketle.mjs";
import { bulguRaporu, derlemKarsilastir, derlemOzeti, kipBelirle, poiKarsilastir, poiOzeti, temelOku } from "./temel.mjs";

const A = argumanlar();
const SINIR_MS = Number(process.env.DERLEM_SINIR_MS ?? 120_000);
const t0 = performance.now();

if (!A["paketleme-yok"]) await paketle();
const dosyalar = derlemDosyalari({ uygulama: A.uygulama, suzgec: A.dosya });
if (!dosyalar.length) {
  console.error("Derlemde dosya bulunamadı (test-derlem/ ve uygulama fikstürleri).");
  process.exit(1);
}
const gercek = {};
for (const p of [join(DERLEM, "gercek_lo.json"), join(URETILEN, "gercek_lo.json")]) if (existsSync(p)) Object.assign(gercek, JSON.parse(readFileSync(p, "utf8")));

const ciktiKok = join(CIKTI, "cikti");
rmSync(ciktiKok, { recursive: true, force: true });
const havuz = new Havuz({ boyut: Math.max(1, Math.min(3, availableParallelism() - 1)), bellekMb: 4096, veri: { paket: PAKET } });
const yiginMi = (h) => h && h.ad === "RangeError" && /call stack/i.test(h.ileti);
const sonuclar = [];
console.log(`${dosyalar.length} dosya turlanıyor…`);

await Promise.all(
  dosyalar.map(async (f) => {
    const onEk = `${f.kaynak}-${createHash("sha1").update(dirname(f.anahtar)).digest("hex").slice(0, 6)}-`;
    const r = await havuz.calistir({ tur: "tur", app: f.app, yol: f.yol, ciktiDizini: join(ciktiKok, f.app), onEk }, SINIR_MS);
    const o = { anahtar: f.anahtar, app: f.app, kaynak: f.kaynak, yol: f.yol, sureMs: r.sureMs };
    if (r.durum !== "tamam") {
      o.durum = r.durum;
      o.havuzHatasi = r.hata;
    } else if (r.cikti.cokme) {
      o.durum = "cokme";
      o.acmaHatasi = r.cikti.cokme;
      o.yigin = yiginMi(r.cikti.cokme);
    } else {
      const c = r.cikti;
      o.durum = "tamam";
      o.boyut = c.boyut;
      o.envanter = c.envanter;
      if (c.acmaHatasi) {
        o.acmaHatasi = c.acmaHatasi;
        o.yigin = yiginMi(c.acmaHatasi);
      } else {
        o.acildi = true;
        Object.assign(o, { bicim: c.bicim, uyari: c.uyari, tAc: c.tAc, tHesap: c.tHesap });
        if (f.app === "tablo") o.loFormul = gercekKarsilastir(gercek[f.anahtar], c.hesap);
        o.turlar = c.turlar.map((tr) => {
          const t = { bicim: tr.bicim, tYaz: tr.tYaz, tYenidenAc: tr.tYenidenAc, ciktiYolu: tr.ciktiYolu, ciktiBoyut: tr.ciktiBoyut };
          if (tr.hata) {
            t.hata = tr.hata;
            t.yigin = yiginMi(tr.hata);
            return t;
          }
          const d = FARK[f.app](c.A, tr.B);
          t.kayip = kayipHaritasi(f.app, d, tr, c.envanter);
          t.kararsiz = tr.C ? kararsizlik(f.app, tr.B, tr.C, tr) : [];
          // Ayrıntı: büyük hücre/stil haritaları rapora girmez
          t.fark = { ...d, ...(f.app === "tablo" ? { etkinStil: tr.etkinStil } : {}) };
          t.envanter = tr.envanter;
          return t;
        });
      }
    }
    const durum = o.durum !== "tamam" ? o.durum.toUpperCase() : o.acildi ? (o.turlar.some((t) => t.hata) ? "tur hatası" : "tamam") : "açılamadı";
    process.stdout.write(`${f.app.padEnd(5)} ${f.anahtar} ${durum} ${r.sureMs} ms\n`);
    sonuclar.push(o);
  }),
);
await havuz.kapat();
sonuclar.sort((a, b) => (a.app + a.anahtar).localeCompare(b.app + b.anahtar));

let commit = "";
try {
  commit = execSync("git rev-parse --short HEAD", { cwd: KOK, stdio: ["ignore", "pipe", "ignore"] }).toString().trim();
} catch {
  /* git yok */
}
const toplamMs = Math.round(performance.now() - t0);
const rapor = { tarih: new Date().toISOString(), commit, sureMs: toplamMs, kapsamli: !A.uygulama && !A.dosya, dosyalar: sonuclar };
mkdirSync(RAPOR, { recursive: true });
writeFileSync(join(RAPOR, "derlem.json"), JSON.stringify(rapor, null, 1));

/* ------------------------------------------------------------ özet (Markdown) */
const L = [];
const p = (s = "") => L.push(s);
const kisa = (o) => Object.entries(o ?? {}).map(([k, v]) => `${k}:${v}`).join(" ");
p(`# Derlem kaydet-aç turu`);
p();
p(`${new Date().toISOString().slice(0, 16).replace("T", " ")} · commit ${commit || "-"} · ${sonuclar.length} dosya · ${sure(toplamMs)}`);
p();
p("| uygulama | dosya | açıldı | açılamadı | ağır (takılma/yığın/bellek) | tur | tur hatası | kayıplı tur | kararsız | kaybolan kelime | LO formül farkı |");
p("|---|---|---|---|---|---|---|---|---|---|---|");
const toplam = {};
for (const app of UYGULAMALAR) {
  const R = sonuclar.filter((r) => r.app === app);
  if (!R.length) continue;
  const turlar = R.flatMap((r) => r.turlar ?? []);
  const satir = {
    dosya: R.length,
    acildi: R.filter((r) => r.acildi).length,
    acilamadi: R.filter((r) => r.durum === "tamam" && !r.acildi).length,
    agir: R.filter((r) => r.durum !== "tamam" || r.yigin || (r.turlar ?? []).some((t) => t.yigin)).length,
    tur: turlar.length,
    turHatasi: turlar.filter((t) => t.hata).length,
    kayipli: turlar.filter((t) => Object.keys(t.kayip ?? {}).length).length,
    kararsiz: turlar.filter((t) => t.kararsiz?.length).length,
    kelime: turlar.reduce((a, t) => a + (t.kayip?.kelime ?? 0), 0),
    lo: R.reduce((a, r) => a + (r.loFormul?.farkli ?? 0), 0),
  };
  toplam[app] = satir;
  p(`| ${app} | ${satir.dosya} | ${satir.acildi} | ${satir.acilamadi} | ${satir.agir} | ${satir.tur} | ${satir.turHatasi} | ${satir.kayipli} | ${satir.kararsiz} | ${satir.kelime} | ${app === "tablo" ? satir.lo : "-"} |`);
}
const poi = poiOzeti(sonuclar);
const temelPoi = A.temelsiz ? {} : (temelOku().poi ?? {});
if (Object.keys(poi).length) {
  p();
  p("## POI derlemi (gerçek Office dosyaları, yerel biçim turu)");
  p();
  p("Kaybolan toplamlar; parantez içinde temel çizgi. Parça kaybı artışı her kipte HATA (temel.json `poi`).");
  p();
  p("| uygulama | dosya | açıldı | parça | ilişki | ns | öğe | model |");
  p("|---|---|---|---|---|---|---|---|");
  const tv = (app, k) => (temelPoi[app] ? ` (${temelPoi[app][k] ?? 0})` : "");
  for (const [app, x] of Object.entries(poi)) p(`| ${app} | ${x.dosya} | ${x.acildi}${tv(app, "acildi")} | ${x.parça}${tv(app, "parça")} | ${x.ilişki}${tv(app, "ilişki")} | ${x.ns}${tv(app, "ns")} | ${x.öğe}${tv(app, "öğe")} | ${x.model}${tv(app, "model")} |`);
}
p();
p("## Kaybolanlar (anahtar başına toplam; tur sayısı)");
p();
p("Model: `kelime`, `düğüm.*`, `işaret.*`, `öznitelik.*`, `ayar.*`, `değer`, `formül`, `stil.*` (etkin stil farkı olan hücre), `yapı.*`, `şekil.*`, `biçim.*`, `slaytAlanı.*`. Paket (aynı biçim ailesinde): `parça:*` (kaybolan ZIP parçası, içerik türüne göre), `ilişki:*` (ilişki türü), `ns:*` (uzantı ad alanı öğesi: w14, x14, p14, mc, dgm, cs…); her ailede `öğe:*` (paragraf, tablo, resim, yorum, dipnot, slayt… sayısındaki azalma).");
p();
const kayipToplam = {};
for (const r of sonuclar)
  for (const t of r.turlar ?? [])
    for (const [k, n] of Object.entries(t.kayip ?? {})) {
      const key = `${r.app} ${k}`;
      kayipToplam[key] ??= { n: 0, tur: 0 };
      kayipToplam[key].n += n;
      kayipToplam[key].tur++;
    }
p("| uygulama ve anahtar | kaybolan | tur |");
p("|---|---|---|");
for (const [k, v] of Object.entries(kayipToplam).sort((a, b) => b[1].tur - a[1].tur || b[1].n - a[1].n)) p(`| ${k} | ${v.n} | ${v.tur} |`);
for (const app of UYGULAMALAR) {
  const R = sonuclar.filter((r) => r.app === app);
  if (!R.length) continue;
  p();
  p(`## ${app}`);
  p();
  p("| dosya | biçim | açılış ms | uyarı / hata | tur | yaz ms | kayıp | 2. tur |");
  p("|---|---|---|---|---|---|---|---|");
  for (const r of R) {
    const ad = r.anahtar.replace(/^(test-derlem|uretilen)\//, "").replace(/^apps\/|^packages\//, "");
    if (!r.acildi) {
      const ileti = r.acmaHatasi?.ileti ?? r.havuzHatasi?.ileti ?? r.durum;
      p(`| ${ad} | - | - | **${r.durum !== "tamam" ? r.durum.toUpperCase() : r.yigin ? "YIĞIN TAŞMASI" : "AÇILAMADI"}** ${String(ileti).replace(/\|/g, "/").slice(0, 120)} | | | | |`);
      continue;
    }
    let bas = `| ${ad} | ${r.bicim} | ${r.tAc} | ${String(r.uyari ?? "").replace(/\|/g, "/").slice(0, 70)}${r.loFormul ? ` formül(LO) ${r.loFormul.n - r.loFormul.farkli}/${r.loFormul.n}` : ""}`;
    for (const t of r.turlar) {
      if (t.hata) {
        p(`${bas} | →${t.bicim} | | **${t.yigin ? "YIĞIN TAŞMASI" : "HATA"}** ${t.hata.ileti.replace(/\|/g, "/").slice(0, 100)} | |`);
      } else {
        p(`${bas} | →${t.bicim} | ${t.tYaz} | ${kisa(t.kayip) || "—"} | ${t.kararsiz.length ? "DEĞİŞİYOR: " + t.kararsiz.slice(0, 6).join(",") : "kararlı"} |`);
      }
      bas = "| | | |";
    }
  }
}
let md = L.join("\n") + "\n";

/* ------------------------------------------------------------ temel çizgi */
let cikisKodu = 0;
if (!A.temelsiz) {
  const temel = temelOku();
  const kip = kipBelirle(A, temel);
  const kapsamli = !A.uygulama && !A.dosya;
  const temelDerlem = kapsamli ? temel.derlem ?? {} : Object.fromEntries(Object.entries(temel.derlem ?? {}).filter(([k]) => sonuclar.some((r) => r.anahtar === k)));
  const b = derlemKarsilastir(derlemOzeti(sonuclar), temelDerlem, kip);
  const { md: kmd, hataVar } = bulguRaporu("Derlem turu", b);
  md += "\n" + kmd + "\n";
  if (hataVar) cikisKodu = 1;
  // POI kapısı yalnız uygulamanın bütün POI dosyaları koşulduysa anlamlıdır (--dosya süzgecinde atlanır).
  if (!A.dosya && Object.keys(poi).length) {
    const pb = poiKarsilastir(poi, temel.poi ?? {}, kip);
    const r = bulguRaporu("POI derlemi kayıp kapısı", pb);
    md += "\n" + r.md + "\n";
    if (r.hataVar) cikisKodu = 1;
  }
}
writeFileSync(join(RAPOR, "derlem.md"), md);
const T = Object.values(toplam).reduce((a, s) => {
  for (const [k, v] of Object.entries(s)) a[k] = (a[k] ?? 0) + v;
  return a;
}, {});
console.log(
  `\nDerlem turu: ${T.dosya} dosya, ${T.acildi} açıldı, ${T.acilamadi} açılamadı, ${T.agir} ağır; ${T.tur} tur, ${T.turHatasi} tur hatası, ${T.kayipli} kayıplı, ${T.kararsiz} kararsız, ${T.kelime} kaybolan kelime (${sure(toplamMs)})`,
);
console.log(`Rapor: ${join(RAPOR, "derlem.md")}  (çıktılar: ${basename(ciktiKok)}/)`);
process.exitCode = cikisKodu;
