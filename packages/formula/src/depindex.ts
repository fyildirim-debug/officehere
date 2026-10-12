/** Bağımlılık dizini: hücre → bağımlı formüller; aralıklar blok ızgarasında tutulur. */
import { walk, type Node } from "./ast";
import { gk, gkLocal, gkSheet, keyCol, keyRow, type RangeDep } from "./engine-types";
import { getFunction } from "./functions/registry";
import { MAX_COLS, MAX_ROWS } from "./references";
import { resolveStructRect, type TableInfo } from "./structref";

export class DependencyIndex {
  private cellDependents = new Map<number, Set<number>>();
  /** Küçük aralıklar: (sütun bloğu, satır bloğu) ızgarası. */
  private gridIndex = new Map<number, Map<number, Set<RangeDep>>>();
  /** Uzun aralıklar: sütun bloğuna göre. */
  private rangeIndex = new Map<number, Map<number, Set<RangeDep>>>();
  private wideRanges = new Map<number, Set<RangeDep>>();

  add(g: number, cellDeps: number[], rangeDeps: RangeDep[]): void {
    for (const d of cellDeps) {
      let set = this.cellDependents.get(d);
      if (!set) this.cellDependents.set(d, (set = new Set()));
      set.add(g);
    }
    for (const rd of rangeDeps) this.addRange(rd);
  }

  remove(g: number, cellDeps: number[], rangeDeps: RangeDep[]): void {
    for (const d of cellDeps) this.cellDependents.get(d)?.delete(g);
    for (const rd of rangeDeps) this.removeRange(rd);
  }

  clear(): void {
    this.cellDependents.clear();
    this.gridIndex.clear();
    this.rangeIndex.clear();
    this.wideRanges.clear();
  }

  /** Konumdaki değişikliğe bağlı formüller. */
  dependentsOf(g: number, out: (f: number) => void): void {
    const set = this.cellDependents.get(g);
    if (set) for (const f of set) out(f);
    const sheet = gkSheet(g);
    const local = gkLocal(g);
    const row = keyRow(local);
    const col = keyCol(local);
    const hit = (rd: RangeDep) => row >= rd.r1 && row <= rd.r2 && col >= rd.c1 && col <= rd.c2;
    const cell = this.gridIndex.get(sheet)?.get((col >> 4) * 16384 + (row >> 6));
    if (cell) for (const rd of cell) if (hit(rd)) out(rd.f);
    const bucket = this.rangeIndex.get(sheet)?.get(col >> 4);
    if (bucket) for (const rd of bucket) if (hit(rd)) out(rd.f);
    const wide = this.wideRanges.get(sheet);
    if (wide) for (const rd of wide) if (hit(rd)) out(rd.f);
  }

  /** Aralığın dizindeki yeri: ızgara blokları, sütun blokları ya da geniş liste. */
  private slots(rd: RangeDep): { kind: "grid" | "col" | "wide"; keys: number[] } {
    const cb1 = rd.c1 >> 4;
    const cb2 = rd.c2 >> 4;
    if (cb2 - cb1 > 32) return { kind: "wide", keys: [] };
    const rb1 = rd.r1 >> 6;
    const rb2 = rd.r2 >> 6;
    const keys: number[] = [];
    if ((cb2 - cb1 + 1) * (rb2 - rb1 + 1) <= 64) {
      for (let cb = cb1; cb <= cb2; cb++) for (let rb = rb1; rb <= rb2; rb++) keys.push(cb * 16384 + rb);
      return { kind: "grid", keys };
    }
    for (let cb = cb1; cb <= cb2; cb++) keys.push(cb);
    return { kind: "col", keys };
  }

  private addRange(rd: RangeDep): void {
    const slot = this.slots(rd);
    if (slot.kind === "wide") {
      let set = this.wideRanges.get(rd.sheet);
      if (!set) this.wideRanges.set(rd.sheet, (set = new Set()));
      set.add(rd);
      return;
    }
    const map = slot.kind === "grid" ? this.gridIndex : this.rangeIndex;
    let idx = map.get(rd.sheet);
    if (!idx) map.set(rd.sheet, (idx = new Map()));
    for (const k of slot.keys) {
      let set = idx.get(k);
      if (!set) idx.set(k, (set = new Set()));
      set.add(rd);
    }
  }

  private removeRange(rd: RangeDep): void {
    const slot = this.slots(rd);
    if (slot.kind === "wide") {
      this.wideRanges.get(rd.sheet)?.delete(rd);
      return;
    }
    const idx = (slot.kind === "grid" ? this.gridIndex : this.rangeIndex).get(rd.sheet);
    if (!idx) return;
    for (const k of slot.keys) {
      const set = idx.get(k);
      if (!set) continue;
      set.delete(rd);
      if (!set.size) idx.delete(k);
    }
  }
}

export interface DepContext {
sheetId(name: string): number | undefined;
definedName(name: string, sheet: number): Node | undefined;
table?(name: string | undefined, sheet: number, row: number, col: number): TableInfo | undefined;
}

/** Formülün başvurduğu hücre/aralıkları, uçucu ve ad kullanımını çıkarır. */
export function analyzeDeps(ast: Node, g: number, sheet: number, ctx: DepContext): { cellDeps: number[]; rangeDeps: RangeDep[]; volatile: boolean; usesNames: boolean } {
  const cellDeps: number[] = [];
  const rangeDeps: RangeDep[] = [];
  let volatile = false;
  let usesNames = false;
  const seenNames = new Set<string>();
  const local = gkLocal(g);
  const addStruct = (ref: { table?: string; items: ("#All" | "#Data" | "#Headers" | "#Totals" | "#This Row")[]; col1?: string; col2?: string }) => {
    const t = ctx.table?.(ref.table, sheet, keyRow(local), keyCol(local));
    if (!t) return;
    const r = resolveStructRect(ref, t, keyRow(local));
    if ("error" in r) return;
    if (r.r1 === r.r2 && r.c1 === r.c2) cellDeps.push(gk(r.sheet, r.r1, r.c1));
    else rangeDeps.push({ sheet: r.sheet, r1: r.r1, c1: r.c1, r2: r.r2, c2: r.c2, f: g });
  };
  const visit = (n: Node, depth: number) => {
    if (n.type === "func") {
      const def = getFunction(n.name);
      if (def?.volatile) volatile = true;
    } else if (n.type === "name") {
      usesNames = true;
      const key = n.name.toLocaleUpperCase("tr-TR");
      if (seenNames.has(key) || depth > 8) return;
      seenNames.add(key);
      const ast = ctx.definedName(n.name, sheet);
      if (ast) walk(ast, (m) => visit(m, depth + 1));
      else addStruct({ items: ["#Data"], table: n.name });
    } else if (n.type === "structref") {
      addStruct(n.ref);
    } else if (n.type === "ref") {
      const t = n.ref;
      if (t.refError) return;
      const sid = t.sheet !== undefined ? ctx.sheetId(t.sheet) : sheet;
      if (sid === undefined) return;
      if (t.kind === "cell") cellDeps.push(gk(sid, t.a.row, t.a.col));
      else {
        const b = t.b!;
        let r1: number, r2: number, c1: number, c2: number;
        if (t.kind === "cols") {
          r1 = 0;
          r2 = MAX_ROWS - 1;
          c1 = Math.min(t.a.col, b.col);
          c2 = Math.max(t.a.col, b.col);
        } else if (t.kind === "rows") {
          c1 = 0;
          c2 = MAX_COLS - 1;
          r1 = Math.min(t.a.row, b.row);
          r2 = Math.max(t.a.row, b.row);
        } else {
          r1 = Math.min(t.a.row, b.row);
          r2 = Math.max(t.a.row, b.row);
          c1 = Math.min(t.a.col, b.col);
          c2 = Math.max(t.a.col, b.col);
        }
        if (r1 === r2 && c1 === c2) cellDeps.push(gk(sid, r1, c1));
        else rangeDeps.push({ sheet: sid, r1, c1, r2, c2, f: g });
      }
    } else if (n.type === "binary" && n.op === ":" && n.left.type === "ref" && n.right.type === "ref") {
      // A1:B2 dışındaki ":" kullanımları (ör. Sayfa1!A1:Sayfa1!B2) — sınırlayıcı kutu
      const l = n.left.ref;
      const r = n.right.ref;
      if (!l.refError && !r.refError && l.kind === "cell" && r.kind === "cell") {
        const sid = l.sheet !== undefined ? ctx.sheetId(l.sheet) : sheet;
        if (sid !== undefined)
          rangeDeps.push({ sheet: sid, r1: Math.min(l.a.row, r.a.row), c1: Math.min(l.a.col, r.a.col), r2: Math.max(l.a.row, r.a.row), c2: Math.max(l.a.col, r.a.col), f: g });
      }
    }
  };
  walk(ast, (n) => visit(n, 0));
  return { cellDeps, rangeDeps, volatile, usesNames };
}
