/** Hesap motoru: yapısal değişiklikler (satır/sütun ekle-sil, kaydırma, taşıma), dizi taşması ve anlık görüntü. */
import type { FormulaEngine } from "./engine";
import { SHEET_STRIDE, cellKey, gk, gkLocal, gkSheet, keyCol, keyRow, safeParse, type CellRecord, type CellSnapshot, type EngineSnapshot, type FormulaInfo, type SheetData } from "./engine-types";
import { MAX_COLS, MAX_ROWS } from "./references";
import { adjustFormula, moveRefsInFormula, removeSheetFromFormula, renameSheetInFormula, type MoveSpec } from "./transform";
import { Errors, isError, type Scalar } from "./values";

/** Satır/sütun ekler (count > 0) ya da siler (count < 0). */
export function structuralChange(e: FormulaEngine, sheet: number, axis: "row" | "col", at: number, count: number): void {
  const s = e.sheetOrThrow(sheet);
  const name = s.name;
  // Taşmaları temizle
  for (const a of [...e.spillAnchors]) {
    const rec = e.sheets.get(gkSheet(a))?.cells.get(gkLocal(a));
    if (rec?.formula) e.clearSpill(a, rec.formula);
  }
  e.spillOwner.clear();
  e.spillAnchors.clear();
  // Hücreleri taşı
  const moved = new Map<number, CellRecord>();
  for (const [k, rec] of s.cells) {
    let r = keyRow(k);
    let c = keyCol(k);
    const idx = axis === "row" ? r : c;
    let ni: number | null;
    if (count > 0) ni = idx >= at ? idx + count : idx;
    else ni = idx < at ? idx : idx < at - count ? null : idx + count;
    if (ni === null) continue;
    if (axis === "row") r = ni;
    else c = ni;
    if (r >= MAX_ROWS || c >= MAX_COLS) continue;
    moved.set(cellKey(r, c), rec);
  }
  s.cells = moved;
  s.extentDirty = true;
  // Gizli satırlar
  if (axis === "row") {
    const shiftSet = (set: Set<number>) => {
      const next = new Set<number>();
      for (const r of set) {
        if (count > 0) next.add(r >= at ? r + count : r);
        else if (r < at) next.add(r);
        else if (r >= at - count) next.add(r + count);
      }
      return next;
    };
    s.hiddenRows = shiftSet(s.hiddenRows);
    s.filteredRows = shiftSet(s.filteredRows);
  }
  // Tüm formülleri güncelle (değişikliğin ötesine başvurmayanlar ayrıştırılmadan atlanır)
  const ch = { sheet: name, axis, at, count };
  for (const id of e.order) {
    const sh = e.sheets.get(id)!;
    for (const [, rec] of sh.cells) {
      if (!rec.formula) continue;
      if (!mayBeAffected(rec.formula, sheet, axis, at)) {
        rec.formula.spill = null;
        rec.formula.spillTarget = null;
        continue;
      }
      const t = adjustFormula(rec.formula.text, sh.name, ch);
      if (t !== rec.formula.text) {
        rec.formula.text = t;
        rec.formula.ast = safeParse(t);
      }
      rec.formula.spill = null;
      rec.formula.spillTarget = null;
    }
  }
  for (const n of e.names.values()) {
    const scopeName = n.sheet !== undefined ? e.getSheetName(n.sheet) ?? name : name;
    const t = adjustFormula(n.text, scopeName, ch);
    if (t !== n.text) {
      n.text = t;
      n.ast = safeParse(t);
    }
  }
  e.reregisterAll();
  e.markAllDirty();
}





/** Formül, yapısal değişiklikten etkilenen bir konuma başvuruyor olabilir mi (bağımlılık listesine göre)? */
function mayBeAffected(info: FormulaInfo, sheet: number, axis: "row" | "col", at: number): boolean {
  if (!info.ast) return true;
  for (const d of info.cellDeps) {
    if (gkSheet(d) !== sheet) continue;
    const local = gkLocal(d);
    if ((axis === "row" ? keyRow(local) : keyCol(local)) >= at) return true;
  }
  for (const rd of info.rangeDeps) if (rd.sheet === sheet && (axis === "row" ? rd.r2 : rd.c2) >= at) return true;
  return false;
}

/**
 * Hücre aralığını yatay/dikey kaydırır (Hücre Ekle/Sil: "hücreleri sağa/aşağı kaydır").
 * Yalnız verilen satır/sütun bandını etkiler.
 */
export function shiftCellsImpl(e: FormulaEngine, sheet: number, r1: number, c1: number, r2: number, c2: number, direction: "down" | "right" | "up" | "left"): void {
  const s = e.sheetOrThrow(sheet);
  const inputs: { row: number; col: number; input: Scalar | { formula: string } }[] = [];
  const h = r2 - r1 + 1;
  const w = c2 - c1 + 1;
  const inBand = (r: number, c: number) => (direction === "down" || direction === "up" ? c >= c1 && c <= c2 && r >= r1 : r >= r1 && r <= r2 && c >= c1);
  const affected: [number, number][] = [];
  for (const [k] of s.cells) {
    const r = keyRow(k);
    const c = keyCol(k);
    if (inBand(r, c)) affected.push([r, c]);
  }
  for (const [r, c] of affected) {
    const rec = s.cells.get(cellKey(r, c))!;
    let nr = r;
    let nc = c;
    if (direction === "down") nr = r + h;
    else if (direction === "right") nc = c + w;
    else if (direction === "up") {
      if (r <= r2) continue;
      nr = r - h;
    } else {
      if (c <= c2) continue;
      nc = c - w;
    }
    inputs.push({ row: nr, col: nc, input: rec.formula ? { formula: rec.formula.text } : rec.value });
  }
  const mv: MoveSpec = { srcSheet: s.name, r1: 0, c1: 0, r2: 0, c2: 0, dRow: 0, dCol: 0, dstSheet: s.name };
  if (direction === "down") Object.assign(mv, { r1, c1, r2: MAX_ROWS - 1, c2, dRow: h });
  else if (direction === "right") Object.assign(mv, { r1, c1, r2, c2: MAX_COLS - 1, dCol: w });
  else if (direction === "up") Object.assign(mv, { r1: r2 + 1, c1, r2: MAX_ROWS - 1, c2, dRow: -h });
  else Object.assign(mv, { r1, c1: c2 + 1, r2, c2: MAX_COLS - 1, dCol: -w });
  for (const [r, c] of affected) e.setValue(sheet, r, c, null);
  for (const it of inputs) {
    if (it.row < MAX_ROWS && it.col < MAX_COLS) e.setInput(sheet, it.row, it.col, it.input);
  }
  rewriteAll(e, (text, sheetName) => moveRefsInFormula(text, sheetName, mv));
}

/** Kes-yapıştır: kaynak dikdörtgeni hedefe taşır, başvuruları günceller. */
export function moveRangeImpl(e: FormulaEngine, srcSheet: number, r1: number, c1: number, r2: number, c2: number, dstSheet: number, dRow: number, dCol: number): void {
  const src = e.sheetOrThrow(srcSheet);
  const dst = e.sheetOrThrow(dstSheet);
  const items: { row: number; col: number; input: Scalar | { formula: string } }[] = [];
  for (let r = r1; r <= r2; r++)
    for (let c = c1; c <= c2; c++) {
      const rec = src.cells.get(cellKey(r, c));
      if (rec) items.push({ row: r + dRow, col: c + dCol, input: rec.formula ? { formula: rec.formula.text } : rec.value });
    }
  // Hedefi temizle, kaynağı temizle, yaz
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) e.setValue(srcSheet, r, c, null);
  for (let r = r1 + dRow; r <= r2 + dRow; r++) for (let c = c1 + dCol; c <= c2 + dCol; c++) if (r >= 0 && c >= 0) e.setValue(dstSheet, r, c, null);
  for (const it of items) if (it.row >= 0 && it.col >= 0) e.setInput(dstSheet, it.row, it.col, it.input);
  // Taşınan formüller kendi içlerindeki göreli başvuruları korur (Excel davranışı)
  const mv: MoveSpec = { srcSheet: src.name, r1, c1, r2, c2, dRow, dCol, dstSheet: dst.name };
  rewriteAll(e, (text, sheetName) => moveRefsInFormula(text, sheetName, mv));
}

function rewriteAll(e: FormulaEngine, fn: (text: string, sheetName: string) => string): void {
  for (const id of e.order) {
    const sh = e.sheets.get(id)!;
    for (const [k, rec] of sh.cells) {
      if (!rec.formula) continue;
      const t = fn(rec.formula.text, sh.name);
      if (t !== rec.formula.text) e.setFormulaInternal(id, keyRow(k), keyCol(k), t);
    }
  }
  for (const n of e.names.values()) {
    const t = fn(n.text, n.sheet !== undefined ? e.getSheetName(n.sheet) ?? "" : "");
    if (t !== n.text) {
      n.text = t;
      n.ast = safeParse(t);
    }
  }
}


export function exportInputsImpl(e: FormulaEngine, withAst = false): EngineSnapshot {
  return {
    sheets: e.order.map((id) => {
      const s = e.sheets.get(id)!;
      const cells: CellSnapshot[] = [];
      for (const [k, rec] of s.cells) {
        if (rec.formula) cells.push(withAst ? { row: keyRow(k), col: keyCol(k), formula: rec.formula.text, ast: rec.formula.ast } : { row: keyRow(k), col: keyCol(k), formula: rec.formula.text });
        else cells.push({ row: keyRow(k), col: keyCol(k), value: rec.value });
      }
      return { id, name: s.name, cells };
    }),
    names: [...e.names.values()].map((n) => ({ name: n.name, text: n.text, sheet: n.sheet })),
  };
}

/** Bir sayfanın girdilerini ve sayfa listesini anlık görüntüden geri yükler. */
export function importInputsImpl(e: FormulaEngine, snap: EngineSnapshot): void {
  const keepIds = new Set(snap.sheets.map((s) => s.id));
  for (const id of [...e.order]) if (!keepIds.has(id)) e.sheets.delete(id);
  e.order = snap.sheets.map((s) => s.id);
  for (const ss of snap.sheets) {
    let s = e.sheets.get(ss.id);
    if (!s) {
      s = { id: ss.id, name: ss.name, cells: new Map(), maxRow: -1, maxCol: -1, extentDirty: false, hiddenRows: new Set(), filteredRows: new Set() };
      e.sheets.set(ss.id, s);
    }
    s.name = ss.name;
    s.cells = new Map();
    s.extentDirty = true;
    for (const c of ss.cells) {
      const k = cellKey(c.row, c.col);
      if (c.formula !== undefined) {
        const ast = c.ast !== undefined ? c.ast : safeParse(c.formula);
        s.cells.set(k, { value: 0, formula: { text: c.formula, ast, volatile: false, usesNames: false, cellDeps: [], rangeDeps: [], spill: null, spillTarget: null } });
      } else s.cells.set(k, { value: c.value ?? null });
    }
    e.nextSheetId = Math.max(e.nextSheetId, ss.id + 1);
  }
  e.names.clear();
  for (const n of snap.names) {
    const u = n.name.toLocaleUpperCase("tr-TR");
    e.names.set(n.sheet !== undefined ? `${n.sheet}!${u}` : u, { name: n.name, text: n.text, ast: safeParse(n.text), sheet: n.sheet });
  }
  e.spillOwner.clear();
  e.spillAnchors.clear();
  e.reregisterAll();
  e.markAllDirty();
}

/** Sayfayı siler; ona başvuran formüller #BAŞV! olur. */
export function removeSheetImpl(e: FormulaEngine, id: number): void {
  const s = e.sheets.get(id);
  if (!s) return;
  for (const [k] of s.cells) e.unregister(id * SHEET_STRIDE + k);
  e.sheets.delete(id);
  e.order = e.order.filter((x) => x !== id);
  for (const sid of e.order) {
    const sh = e.sheets.get(sid)!;
    for (const [k, rec] of sh.cells) {
      if (!rec.formula) continue;
      const t = removeSheetFromFormula(rec.formula.text, s.name);
      if (t !== rec.formula.text) e.setFormulaInternal(sid, keyRow(k), keyCol(k), t);
    }
  }
  for (const [key, n] of e.names) {
    if (n.sheet === id) e.names.delete(key);
    else {
      n.text = removeSheetFromFormula(n.text, s.name);
      n.ast = safeParse(n.text);
    }
  }
  for (const k of [...e.spillOwner.keys()]) if (gkSheet(k) === id) e.spillOwner.delete(k);
  for (const k of [...e.spillAnchors]) if (gkSheet(k) === id) e.spillAnchors.delete(k);
  e.reregisterAll();
  e.markAllDirty();
}

/** Sayfayı yeniden adlandırır ve tüm başvuruları günceller. */
export function renameSheetImpl(e: FormulaEngine, id: number, name: string): void {
  const s = e.sheets.get(id);
  if (!s || s.name === name) return;
  const other = e.getSheetId(name);
  if (other !== undefined && other !== id) throw new Error(`"${name}" adlı sayfa zaten var`);
  const old = s.name;
  s.name = name;
  for (const sid of e.order) {
    const sh = e.sheets.get(sid)!;
    for (const [, rec] of sh.cells) {
      if (!rec.formula) continue;
      const t = renameSheetInFormula(rec.formula.text, old, name);
      if (t !== rec.formula.text) {
        rec.formula.text = t;
        rec.formula.ast = safeParse(t);
      }
    }
  }
  for (const n of e.names.values()) {
    n.text = renameSheetInFormula(n.text, old, name);
    n.ast = safeParse(n.text);
  }
  e.reregisterAll();
  e.markAllDirty();
}

/** Dizi sonucunu komşu hücrelere taşırır; engel varsa #TAŞMA!. */
export function applySpillImpl(e: FormulaEngine, g: number, s: SheetData, info: FormulaInfo, row: number, col: number, data: Scalar[][]): void {
  const rows = data.length;
  const cols = data[0]?.length ?? 0;
  const target = { r1: row, c1: col, r2: row + rows - 1, c2: col + cols - 1 };
  const rec = s.cells.get(cellKey(row, col))!;
  let blocked = target.r2 >= MAX_ROWS || target.c2 >= MAX_COLS;
  if (!blocked) {
    outer: for (let i = 0; i < rows; i++) {
      for (let j = 0; j < cols; j++) {
        if (i === 0 && j === 0) continue;
        const k = cellKey(row + i, col + j);
        if (s.cells.has(k)) {
          blocked = true;
          break outer;
        }
        const owner = e.spillOwner.get(s.id * SHEET_STRIDE + k);
        if (owner !== undefined && owner !== g) {
          blocked = true;
          break outer;
        }
      }
    }
  }
  const old = info.spill;
  const sameShape = old && old.rows === rows && old.cols === cols;
  if (blocked) {
    if (old) e.clearSpill(g, info);
    info.spillTarget = target;
    e.spillAnchors.add(g);
    rec.value = Errors.SPILL;
    return;
  }
  if (!sameShape) {
    if (old) e.clearSpill(g, info);
    for (let i = 0; i < rows; i++)
      for (let j = 0; j < cols; j++) {
        if (i === 0 && j === 0) continue;
        const pk = gk(s.id, row + i, col + j);
        e.spillOwner.set(pk, g);
        e.changed.add(pk);
      }
    e.growExtent(s, target.r2, target.c2);
  } else {
    // Aynı şekil: değişen hücreleri bildir
    for (let i = 0; i < rows; i++)
      for (let j = 0; j < cols; j++) {
        if (i === 0 && j === 0) continue;
        if (!sameScalar(old!.data[i][j], data[i][j])) e.changed.add(gk(s.id, row + i, col + j));
      }
  }
  info.spill = { rows, cols, data };
  info.spillTarget = target;
  e.spillAnchors.add(g);
  rec.value = data[0][0];
}

/** Taşma alanını temizler. */
export function clearSpillImpl(e: FormulaEngine, g: number, info: FormulaInfo): void {
  const sp = info.spill;
  if (sp) {
    const sheet = gkSheet(g);
    const local = gkLocal(g);
    const row = keyRow(local);
    const col = keyCol(local);
    for (let i = 0; i < sp.rows; i++)
      for (let j = 0; j < sp.cols; j++) {
        if (i === 0 && j === 0) continue;
        const pk = gk(sheet, row + i, col + j);
        if (e.spillOwner.get(pk) === g) {
          e.spillOwner.delete(pk);
          e.changed.add(pk);
        }
      }
    const s = e.sheets.get(sheet);
    if (s) s.extentDirty = true;
  }
  info.spill = null;
  info.spillTarget = null;
  e.spillAnchors.delete(g);
}

function sameScalar(a: Scalar, b: Scalar): boolean {
  if (a === b) return true;
  if (isError(a) && isError(b)) return a.code === b.code;
  return false;
}
