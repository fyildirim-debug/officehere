// Kaydet-aç turunun karşılaştırma kuralları.
//
// fark*: özgün açılış (A) ile kaydedilip yeniden açılan (B) model ölçümlerinin farkı (okunur ayrıntı).
// kayipHaritasi: bu farkı ve paket envanteri kaybını düz bir {anahtar: sayı} haritasına çevirir;
// temel çizgi bu haritayı saklar, CI sayının artmasını "yeni kayıp" sayar. Yalnız AZALMALAR kayıptır
// (ör. renk 56→250 eklenen açık biçimdir, kayıp değildir).
// Paket anahtarları (envanter.mjs, yalnız aynı biçim ailesinde): `parça:<içerik türü>`, `ilişki:<ilişki türü>`
// (styles, vbaProject, customXml…), `ns:<önek>:<öğe>` (w14:checkbox, x14:dataValidation, mc:AlternateContent,
// dgm:pt…); her ailede `öğe:<anlamlı öğe>`. Geri kalan anahtarlar model ölçümüdür.
import { envanterFarki, envanterKaybi, GRUPLAR } from "./envanter.mjs";

const kelimeler = (s) => (s ?? "").toLocaleLowerCase("tr").split(/[^\p{L}\p{N}]+/u).filter(Boolean);
/** Kelime torbası karşılaştırması: kaç kelime kayboldu / eklendi. */
export function torba(a, b) {
  const m = new Map();
  let top = 0;
  for (const w of kelimeler(a)) {
    m.set(w, (m.get(w) ?? 0) + 1);
    top++;
  }
  let eklenen = 0;
  for (const w of kelimeler(b)) {
    const c = m.get(w) ?? 0;
    if (c > 0) m.set(w, c - 1);
    else eklenen++;
  }
  let kayip = 0;
  for (const v of m.values()) kayip += v;
  return { top, kayip, eklenen };
}
function ilkFark(a, b) {
  a = (a ?? "").replace(/\s+/g, " ").trim();
  b = (b ?? "").replace(/\s+/g, " ").trim();
  if (a === b) return null;
  let i = 0;
  while (i < a.length && a[i] === b[i]) i++;
  return { konum: i, A: a.slice(Math.max(0, i - 25), i + 45), B: b.slice(Math.max(0, i - 25), i + 45) };
}
function sayiFarki(a = {}, b = {}) {
  const out = {};
  for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) if ((a[k] ?? 0) !== (b[k] ?? 0)) out[k] = [a[k] ?? 0, b[k] ?? 0];
  return out;
}
export const ayniDeger = (x, y) => {
  if (x === y) return true;
  if (typeof x === "boolean") x = x ? 1 : 0;
  if (typeof y === "boolean") y = y ? 1 : 0;
  if (typeof x === "number" && typeof y === "number") return Math.abs(x - y) <= 1e-9 * Math.max(1, Math.abs(x), Math.abs(y));
  if ((x === null && y === "") || (y === null && x === "")) return true;
  return false;
};

function farkBelge(A, B) {
  const d = { kelime: torba(A.text, B.text) };
  const fd = ilkFark(A.text, B.text);
  if (fd) d.metinFarki = fd;
  const n = sayiFarki(A.m.nodes, B.m.nodes);
  const mk = sayiFarki(A.m.marks, B.m.marks);
  const at = sayiFarki(A.m.attrs, B.m.attrs);
  const ay = {};
  for (const k of Object.keys(A.ayar)) if (A.ayar[k] !== B.ayar[k]) ay[k] = [A.ayar[k], B.ayar[k]];
  if (Object.keys(n).length) d.dugum = n;
  if (Object.keys(mk).length) d.isaret = mk;
  if (Object.keys(at).length) d.oznitelik = at;
  if (Object.keys(ay).length) d.ayar = ay;
  if (B.semaTamam !== true) d.sema = B.semaTamam;
  return d;
}

function farkTablo(A, B) {
  const d = { deger: 0, formul: 0, stil: {}, yapi: {}, ornek: [] };
  if (A.sheets.length !== B.sheets.length) d.yapi.sayfaSayısı = [A.sheets.length, B.sheets.length];
  A.sheets.forEach((sa, i) => {
    const sb = B.sheets.find((s) => s.name === sa.name) ?? B.sheets[i];
    if (!sb) return;
    if (sa.name !== sb.name) d.yapi[`ad#${i}`] = [sa.name, sb.name];
    if (sa.cells && sb.cells) {
      for (const k of new Set([...Object.keys(sa.cells), ...Object.keys(sb.cells)])) {
        const [fa, va] = sa.cells[k] ?? [null, null];
        const [fb, vb] = sb.cells[k] ?? [null, null];
        if (!ayniDeger(va, vb)) {
          d.deger++;
          if (d.ornek.length < 6) d.ornek.push(`${sa.name}!${k} değer ${JSON.stringify(va)}→${JSON.stringify(vb)}`);
        }
        if ((fa ?? "") !== (fb ?? "")) {
          d.formul++;
          if (d.ornek.length < 6) d.ornek.push(`${sa.name}!${k} formül ${fa}→${fb}`);
        }
      }
    }
    for (const k of new Set([...Object.keys(sa.styles ?? {}), ...Object.keys(sb.styles ?? {})])) {
      const xa = sa.styles[k] ?? {};
      const xb = sb.styles?.[k] ?? {};
      for (const key of new Set([...Object.keys(xa), ...Object.keys(xb)])) if (String(xa[key] ?? "") !== String(xb[key] ?? "")) d.stil[key] = (d.stil[key] ?? 0) + 1;
    }
    for (const key of ["hidden", "merges", "notes", "links", "validations", "condFormats", "charts", "chartTypes", "images", "shapes", "tables", "pivots", "autoFilter", "frozen", "colWidths", "hiddenCols"]) {
      if (String(sa[key]) !== String(sb[key])) d.yapi[`${sa.name}.${key}`] = [sa[key], sb[key]];
    }
  });
  return d;
}

function farkSunum(A, B) {
  const d = { kelime: torba(A.text, B.text) };
  if (A.slides.length !== B.slides.length) d.slaytSayisi = [A.slides.length, B.slides.length];
  const fd = ilkFark(A.text, B.text);
  if (fd) d.metinFarki = fd;
  const k = sayiFarki(A.kinds, B.kinds);
  if (Object.keys(k).length) d.sekil = k;
  const f = sayiFarki(A.fmt, B.fmt);
  if (Object.keys(f).length) d.bicim = f;
  if (A.tableCells !== B.tableCells) d.tabloHucre = [A.tableCells, B.tableCells];
  const alan = {};
  A.slides.forEach((s, i) => {
    const t = B.slides[i];
    if (!t) return;
    for (const key of ["shapes", "notes", "transition", "anims", "comments", "hidden", "bg"]) if (String(s[key]) !== String(t[key])) alan[key] = (alan[key] ?? 0) + 1;
  });
  if (Object.keys(alan).length) d.slaytAlani = alan;
  if (A.w !== B.w || A.h !== B.h) d.boyut = [`${A.w}x${A.h}`, `${B.w}x${B.h}`];
  return d;
}

export const FARK = { belge: farkBelge, tablo: farkTablo, sunum: farkSunum };

/** Düz kayıp haritası: {anahtar: kaybolan miktar}. Boş harita = kayıpsız. */
export function kayipHaritasi(app, d, tr, kaynakEnvanteri) {
  const k = {};
  const ekle = (key, n) => {
    if (n > 0) k[key] = (k[key] ?? 0) + n;
  };
  const azalma = (grup, onEk) => {
    for (const [key, [a, b]] of Object.entries(grup ?? {})) if (typeof a === "number" && typeof b === "number" && b < a) ekle(`${onEk}.${key}`, a - b);
  };
  if (app === "belge") {
    ekle("kelime", d.kelime?.kayip ?? 0);
    azalma(d.dugum, "düğüm");
    azalma(d.isaret, "işaret");
    azalma(d.oznitelik, "öznitelik");
    for (const key of Object.keys(d.ayar ?? {})) ekle(`ayar.${key}`, 1);
    if (d.sema) ekle("şema", 1);
  } else if (app === "tablo") {
    ekle("değer", d.deger);
    ekle("formül", d.formul);
    for (const [key, n] of Object.entries(tr.etkinStil?.fark ?? {})) ekle(`stil.${key}`, n);
    for (const key of Object.keys(d.yapi ?? {})) ekle(`yapı.${key.replace(/^.*\./, "")}`, 1);
  } else {
    if (d.slaytSayisi && d.slaytSayisi[1] < d.slaytSayisi[0]) ekle("slayt", d.slaytSayisi[0] - d.slaytSayisi[1]);
    ekle("kelime", d.kelime?.kayip ?? 0);
    azalma(d.sekil, "şekil");
    azalma(d.bicim, "biçim");
    if (d.tabloHucre && d.tabloHucre[1] < d.tabloHucre[0]) ekle("tabloHücre", d.tabloHucre[0] - d.tabloHucre[1]);
    for (const [key, n] of Object.entries(d.slaytAlani ?? {})) ekle(`slaytAlanı.${key}`, n);
    if (d.boyut) ekle("slaytBoyutu", 1);
  }
  for (const [key, n] of Object.entries(envanterKaybi(kaynakEnvanteri, tr.envanter))) ekle(key, n);
  return k;
}

/** Kayıp haritasının grup toplamları: {parça, ilişki, ns, öğe, model}. */
export function kayipGruplari(kayip) {
  const g = Object.fromEntries([...GRUPLAR.map(([, onEk]) => [onEk, 0]), ["model", 0]]);
  for (const [k, n] of Object.entries(kayip ?? {})) {
    const onEk = k.slice(0, k.indexOf(":"));
    g[k.includes(":") && onEk in g ? onEk : "model"] += n;
  }
  return g;
}

/** İkinci turda (B→C) değişen alanlar: boş dizi = kararlı. */
export function kararsizlik(app, B, C, tr) {
  const d = FARK[app](B, C);
  const out = [];
  if (app === "tablo") {
    if (d.deger) out.push(`değer:${d.deger}`);
    if (d.formul) out.push(`formül:${d.formul}`);
    for (const key of Object.keys(d.stil)) out.push(`stil.${key}`);
    for (const key of Object.keys(d.yapi)) out.push(`yapı.${key}`);
  } else {
    if (d.metinFarki) out.push("metin");
    for (const key of Object.keys(d)) if (!["kelime", "metinFarki"].includes(key)) out.push(key);
  }
  for (const key of Object.keys(envanterFarki(tr.envanter, tr.envanter2))) out.push(key);
  return out;
}

/** Tablo: motorun hesapladığı formül sonuçlarını LibreOffice'in hesabıyla karşılaştırır. */
export function gercekKarsilastir(gercek, hesap) {
  if (!gercek || gercek.__hata__ || !hesap) return null;
  let n = 0;
  let farkli = 0;
  const ornek = [];
  for (const sh of hesap) {
    const ts = gercek[sh.name];
    if (!ts) continue;
    for (const [k, kayit] of Object.entries(ts)) {
      const tv = Array.isArray(kayit) ? kayit[1] : kayit;
      n++;
      const v = sh.vals[k];
      let ok = ayniDeger(v, tv);
      if (!ok && typeof tv === "string" && typeof v === "string" && tv.startsWith("#") && v.startsWith("#")) ok = tv === v;
      if (!ok && typeof tv === "number" && typeof v === "number" && Math.abs(tv - v) <= 1e-6 * Math.max(1, Math.abs(tv))) ok = true;
      if (!ok) {
        farkli++;
        if (ornek.length < 6) ornek.push(`${sh.name}!${k} LO=${JSON.stringify(tv)} OH=${JSON.stringify(v)}`);
      }
    }
  }
  return { n, farkli, ornek };
}
