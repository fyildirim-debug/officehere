// İşçi iş parçacığı: uygulamanın kendi okuyucu/yazıcısıyla dosyayı açar, ölçer, kaydeder, yeniden açar.
//
// havuz.mjs tarafından worker_threads ile çalıştırılır; süre aşımında ana süreç işçiyi sonlandırır.
// İki iş türü vardır:
//   tur:     aç → ölç (A) → her hedef biçimde yaz → yeniden aç → ölç (B) → bir kez daha yaz-aç (C, yakınsama)
//   bulanik: bozuk dosyayı aç; açılırsa kaydedilebiliyor mu dene
// Ölçümler model düzeyindedir (metin, düğüm/işaret, hücre/formül/stil, slayt/şekil/not) ve ayrıca
// dosyanın paket envanteri (envanter.mjs) çıkarılır. Karşılaştırma ana süreçte (fark.mjs) yapılır.
import "./dom.mjs";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { basename, extname, join } from "node:path";
import { pathToFileURL } from "node:url";
import { parentPort, workerData } from "node:worker_threads";
import { envanter } from "./envanter.mjs";

let B, T, S;
try {
  const yukle = (app) => import(pathToFileURL(join(workerData.paket, `${app}-io.mjs`)).href);
  [B, T, S] = await Promise.all([yukle("belge"), yukle("tablo"), yukle("sunum")]);
} catch (e) {
  parentPort.postMessage({ yuklemeHatasi: String(e?.message ?? e) });
  throw e;
}

const toAB = (u8) => u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
const enc = (d) => (typeof d === "string" ? new TextEncoder().encode(d) : d instanceof Uint8Array ? d : new Uint8Array(d));
const simdi = () => performance.now();
const art = (o, k, n = 1) => (o[k] = (o[k] ?? 0) + n);

/* ------------------------------------------------------------ Belge */
function belgeOlc(doc) {
  // Düzenleyicinin göreceği hâl: şemaya göre normalleştirilmiş JSON.
  let content = doc.content;
  let semaTamam = true;
  try {
    content = B.validateContent(doc.content).toJSON();
  } catch (e) {
    semaTamam = String(e.message).slice(0, 120);
  }
  const m = { nodes: {}, marks: {}, attrs: {} };
  const gez = (n) => {
    art(m.nodes, n.type);
    if (n.type === "text") {
      for (const mk of n.marks ?? []) {
        art(m.marks, mk.type, n.text.length);
        if (mk.type === "textStyle") for (const [k, v] of Object.entries(mk.attrs ?? {})) if (v != null && v !== "" && v !== false) art(m.marks, `textStyle.${k}`, n.text.length);
      }
    }
    const a = n.attrs ?? {};
    if (a.textAlign && a.textAlign !== "left") art(m.attrs, `hiza.${a.textAlign}`);
    if (n.type === "heading") art(m.attrs, `başlık.h${a.level}`);
    if ((a.colspan ?? 1) > 1 || (a.rowspan ?? 1) > 1) art(m.attrs, "birleşikHücre");
    if (n.type === "note") art(m.attrs, `not.${a.kind}`);
    if (a.styleId) art(m.attrs, "stilAtanmış");
    for (const c of n.content ?? []) gez(c);
  };
  gez(content);
  const s = doc.settings ?? {};
  const hf = (j) => (j ? B.docToText(j, "\n").trim().length : 0);
  return {
    semaTamam,
    text: B.docToText(content, "\n"),
    m,
    ayar: {
      üstBilgiKarakter: hf(s.header),
      altBilgiKarakter: hf(s.footer) + (s.footer && JSON.stringify(s.footer).includes("pageField") ? 1000 : 0),
      yorum: Object.keys(s.comments ?? {}).length,
      sayfa: s.page ? `${Math.round(s.page.width ?? 0)}x${Math.round(s.page.height ?? 0)}${s.page.orientation ? "/" + s.page.orientation : ""}` : "",
      başlıkÖzelliği: s.properties?.title ?? "",
    },
  };
}

function belgeYaz(doc, bicim) {
  if (bicim === "odt") return B.writeOdt(doc);
  if (bicim === "rtf") return B.writeRtf(doc);
  if (bicim === "md") return B.writeMarkdown(doc);
  return B.writeDocx(doc);
}

/* ------------------------------------------------------------ Tablo */
const STIL_ANAHTARLARI = ["font", "size", "bold", "italic", "underline", "strike", "color", "fill", "hAlign", "vAlign", "wrap", "rotation", "numFmt", "bTop", "bBottom", "bLeft", "bRight", "indent"];
function normDeger(v) {
  if (v === null || v === undefined) return null;
  if (T.isError(v)) return v.code;
  if (typeof v === "number") return Number.isFinite(v) ? v : String(v);
  return v;
}
function tabloOlc(wb, { hafif = false } = {}) {
  const sheets = wb.sheets.map((sh) => {
    const cells = {};
    let formulas = 0;
    let errors = 0;
    wb.engine.forEachCell(sh.id, (r, c) => {
      const f = wb.getFormula(sh, r, c);
      const v = normDeger(wb.getValue(sh, r, c));
      if (f) formulas++;
      if (typeof v === "string" && v.startsWith("#") && f) errors++;
      cells[`${r},${c}`] = f ? [f, v] : [null, v];
    });
    const styles = {};
    for (const [k, id] of sh.styles) {
      const st = wb.styles.get(id);
      const sig = {};
      for (const key of STIL_ANAHTARLARI) if (st[key] !== undefined && st[key] !== false) sig[key] = typeof st[key] === "object" ? JSON.stringify(st[key]) : st[key];
      if (!hafif) styles[`${T.keyRow(k)},${T.keyCol(k)}`] = sig;
    }
    const sutun = [];
    for (let c = 0; c < 12; c++) sutun.push(Math.round(sh.cols.size(c)));
    let gizliSutun = 0;
    for (let c = 0; c < 30; c++) if (sh.cols.isHidden?.(c)) gizliSutun++;
    return {
      name: sh.name,
      hidden: !!sh.hidden,
      cells: hafif ? undefined : cells,
      nCells: Object.keys(cells).length,
      formulas,
      errors,
      styles,
      merges: sh.merges?.length ?? 0,
      notes: sh.notes?.size ?? 0,
      links: sh.links?.size ?? 0,
      validations: sh.validations?.length ?? 0,
      condFormats: sh.condFormats?.length ?? 0,
      charts: sh.charts?.length ?? 0,
      chartTypes: (sh.charts ?? []).map((c) => c.type ?? c.kind ?? "?").join(","),
      images: sh.images?.length ?? 0,
      shapes: sh.shapes?.length ?? 0,
      tables: sh.tables?.length ?? 0,
      pivots: sh.pivots?.length ?? 0,
      autoFilter: !!sh.autoFilter,
      frozen: `${sh.frozen?.rows ?? 0},${sh.frozen?.cols ?? 0}`,
      colWidths: sutun.join(","),
      hiddenCols: gizliSutun,
    };
  });
  return { sheets };
}

/** İki kitapta etkin (hücre+satır+sütun) stil karşılaştırması: içerikli ya da kendi stili olan tüm hücreler. */
function tabloEtkinStilFarki(wa, wb2) {
  const out = { hücre: 0, fark: {}, örnek: [] };
  wa.sheets.forEach((sa, i) => {
    const sb = wb2.sheets.find((s) => s.name === sa.name) ?? wb2.sheets[i];
    if (!sb) return;
    const keys = new Set();
    for (const [w, s] of [[wa, sa], [wb2, sb]]) {
      w.engine.forEachCell(s.id, (r, c) => keys.add(T.cellKey(r, c)));
      for (const k of s.styles.keys()) keys.add(k);
    }
    for (const k of keys) {
      const r = T.keyRow(k);
      const c = T.keyCol(k);
      // birleşik alanın gizli hücreleri atlanır
      const m = sa.mergeAt?.(r, c);
      if (m && (m.r1 !== r || m.c1 !== c)) continue;
      const xa = wa.styleAt(sa, r, c);
      const xb = wb2.styleAt(sb, r, c);
      out.hücre++;
      for (const key of STIL_ANAHTARLARI) {
        const va = xa[key] === false ? undefined : xa[key];
        const vb = xb[key] === false ? undefined : xb[key];
        const ja = typeof va === "object" ? JSON.stringify(va) : va;
        const jb = typeof vb === "object" ? JSON.stringify(vb) : vb;
        if (String(ja ?? "") !== String(jb ?? "")) {
          out.fark[key] = (out.fark[key] ?? 0) + 1;
          if (out.fark[key] === 1 && out.örnek.length < 8) out.örnek.push(`${sa.name}!R${r + 1}C${c + 1} ${key}: ${ja ?? "-"} → ${jb ?? "-"}`);
        }
      }
    }
  });
  return out;
}

/* ------------------------------------------------------------ Sunum */
function sunumOlc(pres) {
  const kinds = {};
  const fmt = {};
  const text = [];
  let tableCells = 0;
  const duz = (shapes, out = []) => {
    for (const s of shapes) {
      out.push(s);
      if (s.kind === "group") duz(s.children ?? [], out);
    }
    return out;
  };
  const metinOf = (body) => (body?.paragraphs ?? []).map((p) => (p.runs ?? []).map((r) => r.text).join("")).join("\n");
  const slides = pres.slides.map((sl) => {
    const all = duz(sl.shapes ?? []);
    for (const s of all) {
      art(kinds, s.kind);
      if (s.kind === "sp" && s.geom) art(kinds, `sp.${typeof s.geom === "string" ? s.geom : s.geom.preset ?? "özel"}`);
      const bodies = [];
      if (s.text) bodies.push(s.text);
      if (s.kind === "table")
        for (const row of s.rows ?? [])
          for (const cell of row.cells ?? row) {
            tableCells++;
            if (cell?.text) bodies.push(cell.text);
          }
      for (const b of bodies) {
        for (const p of b.paragraphs ?? []) {
          if (p.level > 0) art(fmt, "düzey>0");
          for (const r of p.runs ?? []) {
            if (r.props?.bold) art(fmt, "kalın", r.text.length);
            if (r.props?.italic) art(fmt, "italik", r.text.length);
            if (r.props?.color) art(fmt, "renk", r.text.length);
            if (r.props?.size) art(fmt, "boyut", r.text.length);
            if (r.props?.font) art(fmt, "yazıtipi", r.text.length);
          }
        }
        text.push(metinOf(b));
      }
    }
    return {
      shapes: all.length,
      notes: (sl.notes ?? "").trim(),
      transition: sl.transition?.type ?? "",
      anims: sl.animations?.length ?? 0,
      comments: sl.comments?.length ?? 0,
      hidden: !!sl.hidden,
      bg: sl.background?.type ?? "",
    };
  });
  return { slides, kinds, fmt, tableCells, text: text.join("\n"), w: Math.round(pres.width ?? 0), h: Math.round(pres.height ?? 0) };
}

/* ------------------------------------------------------------ Ortak */
async function ac(app, ad, bytes) {
  const ab = toAB(bytes);
  if (app === "belge") {
    const r = await B.openDocument(ad, ab);
    return { obj: r.doc, bicim: r.format, uyari: r.warning ?? (r.warnings?.length ? r.warnings.join(" ") : undefined) };
  }
  if (app === "tablo") {
    const r = await T.openBook(ad, ab);
    return { obj: r.wb, bicim: r.format, uyari: r.notice ?? r.warning };
  }
  const r = await S.readPresentationFile(new Uint8Array(ab), ad);
  return { obj: r.pres, bicim: r.format, uyari: r.warning ?? (r.warnings?.length ? r.warnings.join(" ") : undefined) };
}
const olc = (app, obj, o) => (app === "belge" ? belgeOlc(obj) : app === "tablo" ? tabloOlc(obj, o) : sunumOlc(obj));
async function yaz(app, obj, bicim) {
  if (app === "belge") return enc(await belgeYaz(obj, bicim));
  if (app === "tablo") return enc(await T.writeBook(obj, obj.active ?? obj.sheets[0], bicim));
  return enc(await S.writePresentationFormat(obj, bicim));
}
/** Kaydedilecek hedef biçimler: her zaman yerel biçim + kaynak biçim yazılabiliyorsa o. */
function hedefler(app, bicim) {
  if (app === "belge") return [...new Set(["docx", ["odt", "rtf", "md"].includes(bicim) ? bicim : "docx"])];
  if (app === "tablo") {
    const w = T.writableAs(bicim);
    return [...new Set(["xlsx", w && w !== "tsv" ? w : "xlsx"])];
  }
  return [...new Set(["pptx", bicim === "odp" ? "odp" : "pptx"])];
}

function hataBilgisi(e) {
  return { ad: e?.constructor?.name ?? typeof e, ileti: String(e?.message ?? e).slice(0, 300), yigin: String(e?.stack ?? "").split("\n").slice(0, 8).join("\n") };
}
const envanterGuvenli = (u8) => {
  try {
    return envanter(u8);
  } catch {
    return null;
  }
};

async function tur(is) {
  const { app, yol, ciktiDizini, onEk } = is;
  const ad = basename(yol);
  const bytes = new Uint8Array(readFileSync(yol));
  const res = { boyut: bytes.length, envanter: envanterGuvenli(bytes) };
  let t = simdi();
  let A;
  try {
    A = await ac(app, ad, bytes);
  } catch (e) {
    res.acmaHatasi = hataBilgisi(e);
    return res;
  }
  res.tAc = Math.round(simdi() - t);
  res.bicim = A.bicim;
  res.uyari = A.uyari;
  res.A = olc(app, A.obj);
  if (app === "tablo") {
    // Motorun kendi hesabı: tam yeniden hesaplama sonrası formül sonuçları (dosyadaki önbellekten bağımsız).
    try {
      const t1 = simdi();
      A.obj.engine.recalc(true);
      res.tHesap = Math.round(simdi() - t1);
      res.hesap = A.obj.sheets.map((sh) => {
        const vals = {};
        A.obj.engine.forEachCell(sh.id, (r, c) => {
          if (A.obj.getFormula(sh, r, c)) vals[`${r},${c}`] = normDeger(A.obj.getValue(sh, r, c));
        });
        return { name: sh.name, vals };
      });
    } catch (e) {
      res.hesapHatasi = hataBilgisi(e);
    }
  }
  res.turlar = [];
  for (const bicim of hedefler(app, A.bicim)) {
    const tr = { bicim };
    try {
      t = simdi();
      const out = await yaz(app, A.obj, bicim);
      tr.tYaz = Math.round(simdi() - t);
      tr.ciktiBoyut = out.length;
      mkdirSync(ciktiDizini, { recursive: true });
      const ciktiYolu = join(ciktiDizini, `${onEk}${basename(ad, extname(ad))}__${extname(ad).slice(1)}.${bicim}`);
      writeFileSync(ciktiYolu, out);
      tr.ciktiYolu = ciktiYolu;
      tr.envanter = envanterGuvenli(out);
      t = simdi();
      const Bk = await ac(app, basename(ciktiYolu), out);
      tr.tYenidenAc = Math.round(simdi() - t);
      tr.B = olc(app, Bk.obj);
      if (app === "tablo") tr.etkinStil = tabloEtkinStilFarki(A.obj, Bk.obj);
      // İkinci tur: yakınsama (aynı biçimde bir kez daha yaz-aç)
      const out2 = await yaz(app, Bk.obj, bicim);
      tr.envanter2 = envanterGuvenli(out2);
      const C = await ac(app, basename(ciktiYolu), out2);
      tr.C = olc(app, C.obj);
    } catch (e) {
      tr.hata = hataBilgisi(e);
    }
    res.turlar.push(tr);
  }
  return res;
}

async function bulanik(is) {
  const { app, yol, ad } = is;
  const bytes = new Uint8Array(readFileSync(yol));
  const t = simdi();
  try {
    const r = await ac(app, ad, bytes);
    const ms = Math.round(simdi() - t);
    // Açılan nesne kaydedilebiliyor mu (bozuk girdi yazıcıyı çökertmesin)
    let yazmaHatasi = null;
    const t2 = simdi();
    try {
      await yaz(app, r.obj, app === "belge" ? "docx" : app === "tablo" ? "xlsx" : "pptx");
    } catch (e) {
      yazmaHatasi = hataBilgisi(e);
    }
    let ozet = "";
    try {
      const m = olc(app, r.obj, { hafif: true });
      ozet =
        app === "belge"
          ? `metin=${m.text.length} şema=${m.semaTamam === true ? "tamam" : "HATA"}`
          : app === "tablo"
            ? `sayfa=${m.sheets.length} hücre=${m.sheets.reduce((a, s) => a + s.nCells, 0)}`
            : `slayt=${m.slides.length}`;
    } catch (e) {
      ozet = `ölçüm hatası: ${e.message}`;
    }
    return { acildi: true, ms, bicim: r.bicim, uyari: r.uyari, yazmaHatasi, tYaz: Math.round(simdi() - t2), ozet };
  } catch (e) {
    return { acildi: false, ms: Math.round(simdi() - t), hata: hataBilgisi(e) };
  }
}

parentPort.on("message", async (is) => {
  try {
    const cikti = is.tur === "bulanik" ? await bulanik(is) : await tur(is);
    parentPort.postMessage({ id: is.id, cikti });
  } catch (e) {
    parentPort.postMessage({ id: is.id, cikti: { cokme: hataBilgisi(e) } });
  }
});
parentPort.postMessage({ hazir: true });
