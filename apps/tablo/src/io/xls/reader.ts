/**
 * Excel 97-2003 ikili çalışma kitabı (.xls, BIFF8; BIFF5 için temel destek)
 * okuyucusu. Bileşik dosyadaki "Workbook"/"Book" akışı kayıt kayıt okunur:
 * sayfalar, paylaşılan dizgiler (SST + CONTINUE), hücre değerleri, formüller
 * (ptg → A1; çevrilemezse önbellek değeri), biçimler, birleştirmeler, boyutlar.
 * Yalnız okunur; kaydederken .xlsx yazılır.
 */
import { isCfb, readCfb } from "@officehere/core";
import { cellKey, errorFromCode, isDateFormat, MAX_COLS, MAX_ROWS, type Scalar } from "@officehere/formula";
import { ptToPx, widthCharsToPx } from "../../model/axis";
import { DEFAULT_FONT, DEFAULT_SIZE, type CellStyle } from "../../model/styles";
import { activateTheme } from "../../model/themeRuntime";
import { WorkbookModel, type CellInput, type SheetModel } from "../../model/workbook";
import { CELL_BUDGET, fitMerges, FormatError, safeSheetName } from "../bookUtil";
import { decodeFormula, errorText, sharedRef, type XlsFormulaContext } from "./formula";
import { byteDecoder, ContinuedReader, f64, joined, R, readByteString, readRecords, readString, rk, u16, u32, u8, type BiffRecord } from "./records";
import { XlsStyles } from "./styles";

/** 1904 tarih sistemindeki seri numarasının 1900 sistemindeki karşılığı için eklenen gün. */
const DAYS_1904 = 1462;

interface Bound {
  offset: number;
  hidden: boolean;
  type: number;
  name: string;
}

interface PendingFormula {
  row: number;
  col: number;
  anchor: { row: number; col: number };
  cached: Scalar;
}

/** Bileşik dosyadan çalışma kitabını okur. */
export function readXls(data: ArrayBuffer | Uint8Array): WorkbookModel {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (!isCfb(bytes)) throw new FormatError("Dosya geçerli bir 97-2003 çalışma kitabı değil.");
  let cfb;
  try {
    cfb = readCfb(bytes);
  } catch (e) {
    throw new FormatError(`Dosya bozuk görünüyor (${(e as Error).message}).`);
  }
  const stream = cfb.read("Workbook") ?? cfb.read("Book");
  if (!stream) {
    if (cfb.has("WordDocument")) throw new FormatError("Bu bir belge dosyası; OfficeHere Belge ile açın.");
    if (cfb.has("PowerPoint Document")) throw new FormatError("Bu bir sunu dosyası; OfficeHere Sunum ile açın.");
    if (cfb.has("EncryptedPackage")) throw new FormatError("Bu dosya parola ile şifrelenmiş. Şifreli dosyalar açılamaz.");
    throw new FormatError("Dosyada çalışma kitabı verisi bulunamadı.");
  }
  return readBiffStream(stream);
}

/** Workbook akışını çözer (testler için ayrıca dışa açık). */
export function readBiffStream(stream: Uint8Array): WorkbookModel {
  const recs = readRecords(stream);
  const bof = recs[0];
  if (!bof || (bof.type !== R.BOF && bof.type !== R.BOF_OLD)) throw new FormatError("Dosya geçerli bir 97-2003 çalışma kitabı değil (BOF kaydı yok).");
  const vers = u16(bof.data, 0);
  if (vers !== 0x0600 && vers !== 0x0500) throw new FormatError("Bu çalışma kitabı çok eski bir biçimde (BIFF2-4) kaydedilmiş; açılamıyor.");
  const biff: 5 | 8 = vers === 0x0600 ? 8 : 5;
  const warnings = new Set<string>();
  let decode = byteDecoder(1252);
  const styles = new XlsStyles(biff, (b) => decode(b));
  const sst: string[] = [];
  const bounds: Bound[] = [];
  let date1904 = false;
  let activeTab = 0;
  const ctx: XlsFormulaContext = { sheetNames: [], xti: [], supbooks: [], names: [] };
  const nameRecs: { name: string; builtin: boolean; code: number; itab: number; rgce: Uint8Array; extra: Uint8Array; hidden: boolean; func: boolean }[] = [];
  // Genel kayıtlar (ilk EOF'a kadar)
  let i = 1;
  for (; i < recs.length; i++) {
    const r = recs[i];
    const d = r.data;
    if (r.type === R.EOF) break;
    switch (r.type) {
      case R.FILEPASS:
        throw new FormatError("Bu dosya parola ile korunuyor (şifreli). Şifreli 97-2003 çalışma kitapları açılamaz; dosyayı parolasız kaydedip yeniden deneyin.");
      case R.CODEPAGE:
        decode = byteDecoder(u16(d, 0));
        break;
      case R.DATEMODE:
        date1904 = u16(d, 0) === 1;
        break;
      case R.WINDOW1:
        activeTab = u16(d, 10);
        break;
      case R.BOUNDSHEET: {
        const name = biff === 8 ? readString(d, 6, 1).text : readByteString(d, 6, 1, decode).text;
        bounds.push({ offset: u32(d, 0), hidden: u8(d, 4) !== 0, type: u8(d, 5), name });
        break;
      }
      case R.FONT:
        styles.font(d);
        break;
      case R.FORMAT:
        styles.format(d);
        break;
      case R.XF:
        styles.xf(d);
        break;
      case R.PALETTE:
        styles.setPalette(d);
        break;
      case R.SST: {
        const rd = new ContinuedReader(r, 8);
        const unique = Math.min(u32(d, 4), 10_000_000);
        for (let k = 0; k < unique && !rd.eof; k++) sst.push(rd.string(2));
        break;
      }
      case R.SUPBOOK: {
        const cch = u16(d, 2);
        ctx.supbooks.push({ kind: cch === 0x0401 ? "self" : cch === 0x3a01 ? "addin" : "external", externNames: [] });
        break;
      }
      case R.EXTERNNAME: {
        const sb = ctx.supbooks[ctx.supbooks.length - 1];
        if (sb && biff === 8) sb.externNames.push(readString(d, 6, 1).text);
        break;
      }
      case R.EXTERNSHEET: {
        if (biff !== 8) break;
        const all = joined(r);
        const n = u16(all, 0);
        for (let k = 0; k < n; k++) ctx.xti.push({ sup: u16(all, 2 + k * 6), first: u16(all, 4 + k * 6), last: u16(all, 6 + k * 6) });
        break;
      }
      case R.NAME: {
        if (biff !== 8) break;
        const all = joined(r);
        const grbit = u16(all, 0);
        const cch = u8(all, 3);
        const cce = u16(all, 4);
        const itab = u16(all, 8);
        // Ad: uzunluk 3. baytta, seçenek baytı 14'te, karakterler ardından
        const high = u8(all, 14) & 1;
        const nameStart = 15;
        const nameEnd = nameStart + (high ? cch * 2 : cch);
        const builtin = !!(grbit & 0x20);
        let name = "";
        for (let k = 0; k < cch; k++) name += String.fromCharCode(high ? u16(all, nameStart + k * 2) : u8(all, nameStart + k));
        const code = builtin ? name.charCodeAt(0) : -1;
        nameRecs.push({ name, builtin, code, itab, rgce: all.subarray(nameEnd, nameEnd + cce), extra: all.subarray(nameEnd + cce), hidden: !!(grbit & 1), func: !!(grbit & 0x02) });
        ctx.names.push({ name, builtin });
        break;
      }
    }
  }
  const wb = new WorkbookModel(true);
  // Varsayılan yazı tipi tema gövde yazı tipi olur
  const font0 = styles.defaultFont();
  if (font0?.name && font0.name !== DEFAULT_FONT) wb.theme = { ...wb.theme, minorFont: font0.name };
  activateTheme(wb.theme);
  const styleIds = new Map<number, number>();
  const styleId = (ixfe: number): number => {
    let id = styleIds.get(ixfe);
    if (id === undefined) {
      const st: CellStyle = styles.style(ixfe);
      if (st.size === DEFAULT_SIZE) delete st.size;
      id = wb.styles.intern(st);
      styleIds.set(ixfe, id);
    }
    return id;
  };
  const baseId = styleId(15);
  // Sayfalar
  const byOffset = new Map<number, number>();
  recs.forEach((r, k) => byOffset.set(r.offset, k));
  const sheetOf = new Map<number, SheetModel>();
  ctx.sheetNames = bounds.map((b) => b.name);
  const rename = new Map<string, string>();
  bounds.forEach((b, k) => {
    if (b.type !== 0) {
      warnings.add(b.type === 2 ? "grafik sayfaları" : "makro sayfaları");
      return;
    }
    const sheet = wb.addSheet(safeSheetName(wb, b.name));
    rename.set(b.name, sheet.name);
    if (b.hidden) sheet.hidden = true;
    sheetOf.set(k, sheet);
  });
  // Formüllerde sayfa adları modeldeki adlarla
  ctx.sheetNames = bounds.map((b) => rename.get(b.name) ?? b.name);
  let budget = CELL_BUDGET;
  bounds.forEach((b, k) => {
    const sheet = sheetOf.get(k);
    const start = byOffset.get(b.offset);
    if (!sheet || start === undefined) return;
    try {
      budget = readSheet(recs, start, sheet, { wb, biff, sst, styleId, baseId, ctx, decode, date1904, warnings, budget });
    } catch {
      warnings.add(`"${sheet.name}" sayfasının bir bölümü`);
    }
  });
  // Adlar (yerleşik: yazdırma alanı ve otomatik filtre)
  for (const n of nameRecs) {
    const scope = n.itab ? sheetOf.get(n.itab - 1) : undefined;
    const f = decodeFormula(n.rgce, n.extra, ctx, { row: 0, col: 0 });
    if (!f || n.func) continue;
    if (n.builtin) {
      const target = scope;
      const m = /^=(?:'(?:[^']|'')+'|[^!]+)!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/.exec(f);
      if (target && m) {
        const rect = { r1: Number(m[2]) - 1, c1: colIndex(m[1]), r2: Number(m[4]) - 1, c2: colIndex(m[3]) };
        if (n.code === 0x0d) target.autoFilter = { range: rect, columns: {} };
        else if (n.code === 0x06) target.page.printArea = rect;
      }
      continue;
    }
    if (n.name.startsWith("_xlfn.")) continue;
    try {
      wb.engine.defineName(n.name, f, scope?.id);
    } catch {
      /* geçersiz ad */
    }
  }
  if (!wb.sheets.length) wb.addSheet("Sayfa1");
  const act = sheetOf.get(activeTab);
  wb.activeIndex = Math.max(0, act && !act.hidden ? wb.sheets.indexOf(act) : wb.sheets.findIndex((s) => !s.hidden));
  wb.syncTables();
  wb.engine.recalc(true);
  for (const s of wb.sheets) {
    const note = fitMerges(wb, s);
    if (note) warnings.add(note);
  }
  wb.loadWarnings = [...warnings];
  return wb;
}

function colIndex(name: string): number {
  let c = 0;
  for (const ch of name) c = c * 26 + (ch.charCodeAt(0) - 64);
  return c - 1;
}

interface SheetCtx {
  wb: WorkbookModel;
  biff: 5 | 8;
  sst: string[];
  styleId: (ixfe: number) => number;
  baseId: number;
  ctx: XlsFormulaContext;
  decode: (b: Uint8Array) => string;
  date1904: boolean;
  warnings: Set<string>;
  budget: number;
}

/** Bir çalışma sayfası alt akışını okur; kalan hücre bütçesini döndürür. */
function readSheet(recs: BiffRecord[], start: number, sheet: SheetModel, c: SheetCtx): number {
  const { wb } = c;
  let budget = c.budget;
  const shared = new Map<string, { r1: number; r2: number; c1: number; c2: number; rgce: Uint8Array; extra: Uint8Array }>();
  const arrays = new Set<string>();
  const pending: PendingFormula[] = [];
  let lastFormula: { row: number; col: number; pend?: PendingFormula } | null = null;
  let frozen = false;
  const put = (row: number, col: number, ixfe: number, v: CellInput | undefined) => {
    if (row >= MAX_ROWS || col >= MAX_COLS || budget <= 0) return;
    budget--;
    const id = c.styleId(ixfe);
    if (id) sheet.styles.set(cellKey(row, col), id);
    if (v === undefined || v === null) return;
    // 1904 tarih sistemi: tarih biçimli sayılar 1900 sistemine kaydırılır
    if (c.date1904 && typeof v === "number" && isDateFormat(wb.styles.get(id).numFmt)) v = v + DAYS_1904;
    wb.setInputRaw(sheet, row, col, v);
  };
  const setFormula = (row: number, col: number, ixfe: number, f: string) => {
    if (row >= MAX_ROWS || col >= MAX_COLS) return;
    const id = c.styleId(ixfe);
    if (id) sheet.styles.set(cellKey(row, col), id);
    wb.setInputRaw(sheet, row, col, { formula: f });
  };
  let depth = 0;
  for (let i = start + 1; i < recs.length; i++) {
    const r = recs[i];
    const d = r.data;
    if (r.type === R.BOF || r.type === R.BOF_OLD) {
      // Gömülü grafik alt akışı
      depth++;
      c.warnings.add("grafikler");
      continue;
    }
    if (r.type === R.EOF) {
      if (depth === 0) break;
      depth--;
      continue;
    }
    if (depth) continue;
    switch (r.type) {
      case R.LABELSST:
        put(u16(d, 0), u16(d, 2), u16(d, 4), c.sst[u32(d, 6)] ?? "");
        break;
      case R.LABEL:
      case R.RSTRING:
        put(u16(d, 0), u16(d, 2), u16(d, 4), c.biff === 8 ? new ContinuedReader(r, 6).string(2) : readByteString(d, 6, 2, c.decode).text);
        break;
      case R.NUMBER:
        put(u16(d, 0), u16(d, 2), u16(d, 4), f64(d, 6));
        break;
      case R.RK:
        put(u16(d, 0), u16(d, 2), u16(d, 4), rk(u32(d, 6)));
        break;
      case R.MULRK: {
        const row = u16(d, 0);
        const first = u16(d, 2);
        const n = Math.floor((d.length - 6) / 6);
        for (let k = 0; k < n; k++) put(row, first + k, u16(d, 4 + k * 6), rk(u32(d, 6 + k * 6)));
        break;
      }
      case R.BOOLERR: {
        const v = u8(d, 6);
        put(u16(d, 0), u16(d, 2), u16(d, 4), u8(d, 7) ? (errorFromCode(errorText(v)) ?? null) : v !== 0);
        break;
      }
      case R.BLANK:
        put(u16(d, 0), u16(d, 2), u16(d, 4), undefined);
        break;
      case R.MULBLANK: {
        const row = u16(d, 0);
        const first = u16(d, 2);
        const n = Math.floor((d.length - 6) / 2);
        for (let k = 0; k < n; k++) put(row, first + k, u16(d, 4 + k * 2), undefined);
        break;
      }
      case R.FORMULA: {
        const all = joined(r);
        const row = u16(all, 0);
        const col = u16(all, 2);
        const ixfe = u16(all, 4);
        let cached: Scalar = null;
        if (u16(all, 12) === 0xffff) {
          const t = u8(all, 6);
          cached = t === 1 ? u8(all, 8) !== 0 : t === 2 ? (errorFromCode(errorText(u8(all, 8))) ?? null) : t === 3 ? "" : null;
        } else cached = f64(all, 6);
        const cce = u16(all, 20);
        const rgce = all.subarray(22, 22 + cce);
        const extra = all.subarray(22 + cce);
        lastFormula = { row, col };
        const anchor = c.biff === 8 ? sharedRef(rgce) : null;
        if (anchor) {
          const p: PendingFormula = { row, col, anchor, cached };
          pending.push(p);
          lastFormula.pend = p;
          put(row, col, ixfe, cached);
          break;
        }
        const f = c.biff === 8 ? decodeFormula(rgce, extra, c.ctx, { row, col }) : undefined;
        if (f) setFormula(row, col, ixfe, f);
        else {
          c.warnings.add(c.biff === 8 ? "çevrilemeyen formüller (son hesaplanan değerleri kullanıldı)" : "formüller (95 biçiminde yalnız son hesaplanan değerleri alınır)");
          put(row, col, ixfe, cached);
        }
        break;
      }
      case R.STRING: {
        // Önceki formülün metin sonucu
        const s = c.biff === 8 ? new ContinuedReader(r, 0).string(2) : readByteString(d, 0, 2, c.decode).text;
        if (lastFormula) {
          if (lastFormula.pend) lastFormula.pend.cached = s;
          if (!wb.getFormula(sheet, lastFormula.row, lastFormula.col)) wb.setInputRaw(sheet, lastFormula.row, lastFormula.col, s);
        }
        break;
      }
      case R.SHRFMLA: {
        const all = joined(r);
        const cce = u16(all, 8);
        shared.set(`${u16(all, 0)},${u8(all, 4)}`, { r1: u16(all, 0), r2: u16(all, 2), c1: u8(all, 4), c2: u8(all, 5), rgce: all.subarray(10, 10 + cce), extra: all.subarray(10 + cce) });
        break;
      }
      case R.ARRAY:
        arrays.add(`${u16(d, 0)},${u8(d, 4)}`);
        break;
      case R.MERGEDCELLS: {
        const n = u16(d, 0);
        for (let k = 0; k < n; k++) {
          const o = 2 + k * 8;
          if (o + 8 > d.length) break;
          const m = { r1: u16(d, o), r2: u16(d, o + 2), c1: u16(d, o + 4), c2: u16(d, o + 6) };
          if (m.r2 >= m.r1 && m.c2 >= m.c1 && (m.r2 > m.r1 || m.c2 > m.c1)) sheet.merges.push(m);
        }
        break;
      }
      case R.COLINFO: {
        const first = u16(d, 0);
        const last = Math.min(u16(d, 2), 255);
        const px = widthCharsToPx(u16(d, 4) / 256);
        const ixfe = u16(d, 6);
        const flags = u16(d, 8);
        const id = c.styleId(ixfe);
        for (let col = first; col <= last; col++) {
          if (Math.abs(px - sheet.cols.defaultSize) > 0.5) sheet.cols.set(col, px);
          if (flags & 0x01) sheet.cols.setHidden(col, true);
          const lvl = (flags >> 8) & 0x07;
          if (lvl) sheet.cols.setLevel(col, lvl);
          if (id && id !== c.baseId) sheet.colStyles.set(col, id);
        }
        break;
      }
      case R.DEFCOLWIDTH: {
        const px = widthCharsToPx(u16(d, 0) + 0.43);
        if (Math.abs(px - sheet.cols.defaultSize) > 1) sheet.cols.setDefault(px);
        break;
      }
      case R.DEFAULTROWHEIGHT: {
        const px = ptToPx(u16(d, 2) / 20);
        if (px > 0 && Math.abs(px - sheet.rows.defaultSize) > 0.5) sheet.rows.setDefault(px);
        break;
      }
      case R.ROW: {
        const row = u16(d, 0);
        if (row >= MAX_ROWS) break;
        const tw = u16(d, 6) & 0x7fff;
        const flags = u32(d, 12);
        const custom = !!(flags & 0x40);
        const px = ptToPx(tw / 20);
        if (tw && (custom || Math.abs(px - sheet.rows.defaultSize) > 0.5)) sheet.rows.set(row, px, custom);
        if (flags & 0x20) sheet.rows.setHidden(row, true);
        if (flags & 0x07) sheet.rows.setLevel(row, flags & 0x07);
        if (flags & 0x10) sheet.rows.setCollapsed(row, true);
        if (flags & 0x80) {
          const id = c.styleId((flags >>> 16) & 0x0fff);
          if (id && id !== c.baseId) sheet.rowStyles.set(row, id);
        }
        break;
      }
      case R.WINDOW2: {
        const g = u16(d, 0);
        if (!(g & 0x02)) sheet.showGridlines = false;
        if (!(g & 0x04)) sheet.showHeadings = false;
        frozen = !!(g & 0x08);
        break;
      }
      case R.PANE:
        if (frozen) sheet.frozen = { rows: u16(d, 2), cols: u16(d, 0) };
        break;
      case R.SCL: {
        const z = Math.round((u16(d, 0) / Math.max(1, u16(d, 2))) * 100);
        if (z >= 10 && z <= 400) sheet.zoom = z;
        break;
      }
      case 0x001c:
        c.warnings.add("notlar");
        break;
      case R.HLINK:
        c.warnings.add("köprüler");
        break;
      case 0x005d: {
        const ot = u16(d, 4);
        if (u16(d, 0) === 0x15 && (ot === 0x05 || ot === 0x08)) c.warnings.add(ot === 0x05 ? "grafikler" : "resimler");
        break;
      }
    }
  }
  // Paylaşılan formüller: her hücre kendi konumuna göre çözülür
  for (const p of pending) {
    const sh = shared.get(`${p.anchor.row},${p.anchor.col}`);
    const inRange = sh && p.row >= sh.r1 && p.row <= sh.r2 && p.col >= sh.c1 && p.col <= sh.c2;
    const f = inRange ? decodeFormula(sh.rgce, sh.extra, c.ctx, { row: p.row, col: p.col }) : undefined;
    if (f) wb.setInputRaw(sheet, p.row, p.col, { formula: f });
    else if (!arrays.has(`${p.anchor.row},${p.anchor.col}`)) c.warnings.add("çevrilemeyen formüller (son hesaplanan değerleri kullanıldı)");
  }
  return budget;
}
