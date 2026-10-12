/**
 * PivotTable hesap çekirdeği (DOM'suz, birim testli): kaynak aralığı
 * kayıtlara çevirir, filtreler, satır/sütun ağaçlarını kurar, değerleri
 * özetler ve Sıkıştırılmış / Anahat / Tablo düzeninde hücre listesi üretir.
 */
import { compareText, Errors, isError, numEquals, type Scalar } from "@officehere/formula";
import type { PivotDef, PivotFn, PivotValueField } from "../model/types";

export interface PivotSource {
  headers: string[];
  records: Scalar[][];
  /** Kaynak sütunlarının sayı biçimleri (ilk veri satırından). */
  numFmts: (string | undefined)[];
}

export type PivotCellKind =
  | "filterLabel"
  | "filterValue"
  | "caption"
  | "colLabelsHeader"
  | "rowLabelsHeader"
  | "fieldHeader"
  | "colItem"
  | "rowItem"
  | "rowParent"
  | "value"
  | "subtotalLabel"
  | "subtotalValue"
  | "grandLabel"
  | "grandValue"
  | "grandColHeader"
  | "blank";

export interface PivotCell {
  row: number;
  col: number;
  value: Scalar;
  kind: PivotCellKind;
  indent?: number;
  numFmt?: string;
  /** Bu hücre bir "grup" satırında mı (kalın satır). */
  bold?: boolean;
  /** Şeritli satırlar için veri satırı sırası. */
  band?: number;
}

export interface PivotLayout {
  rows: number;
  cols: number;
  cells: PivotCell[];
  /** Tablonun (filtreler hariç) başladığı satır ofseti. */
  tableRow: number;
  /** Başlık satırı sayısı (tablonun içinde). */
  headerRows: number;
  /** Etiket sütunu sayısı. */
  labelCols: number;
  empty: boolean;
}

export const FN_LABEL: Record<PivotFn, string> = {
  sum: "Topla",
  count: "Say",
  average: "Ortalama",
  max: "En Büyük",
  min: "En Küçük",
  product: "Çarpım",
  countNums: "Sayıları Say",
};

const FN_CAPTION: Record<PivotFn, string> = {
  sum: "Toplam",
  count: "Say",
  average: "Ortalama",
  max: "En Büyük",
  min: "En Küçük",
  product: "Çarpım",
  countNums: "Sayıları Say",
};

export const BLANK_ITEM = "(boş)";
export const VALUES_FIELD = -1;

export function valueCaption(vf: PivotValueField, headers: string[]): string {
  return vf.name ?? `${FN_CAPTION[vf.fn]} ${headers[vf.field] ?? ""}`;
}

/** Öğe anahtarı: büyük/küçük harf duyarsız metin, sayılar sayısal. */
export function itemKey(v: Scalar): string {
  if (v === null || v === "") return "\u0000b";
  if (typeof v === "number") return `n${v}`;
  if (typeof v === "boolean") return v ? "bT" : "bF";
  if (isError(v)) return `e${v.code}`;
  return `s${String(v).toLocaleUpperCase("tr-TR")}`;
}

export function itemLabel(v: Scalar): Scalar {
  return v === null || v === "" ? BLANK_ITEM : v;
}

function compareItems(a: Scalar, b: Scalar): number {
  const rank = (v: Scalar) => (v === null || v === "" ? 9 : typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 3);
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === "number" && typeof b === "number") return numEquals(a, b) ? 0 : a < b ? -1 : 1;
  if (typeof a === "string" && typeof b === "string") return compareText(a, b);
  if (typeof a === "boolean" && typeof b === "boolean") return a === b ? 0 : a ? 1 : -1;
  return 0;
}

/** Bir alanın benzersiz öğeleri (sıralı, gösterim değeriyle). */
export function fieldItems(src: PivotSource, field: number): { key: string; value: Scalar }[] {
  const seen = new Map<string, Scalar>();
  for (const r of src.records) {
    const v = r[field] ?? null;
    const k = itemKey(v);
    if (!seen.has(k)) seen.set(k, v);
  }
  return [...seen.entries()].map(([key, value]) => ({ key, value })).sort((a, b) => compareItems(a.value, b.value));
}

/* ----------------------------- Özetleme ----------------------------- */

class Acc {
  sum = 0;
  count = 0;
  nums = 0;
  min = Infinity;
  max = -Infinity;
  product = 1;
  error: Scalar = null;

  add(v: Scalar): void {
    if (v === null || v === "") return;
    this.count++;
    if (isError(v)) {
      if (this.error === null) this.error = v;
      return;
    }
    if (typeof v !== "number") return;
    this.nums++;
    this.sum += v;
    this.product *= v;
    if (v < this.min) this.min = v;
    if (v > this.max) this.max = v;
  }

  result(fn: PivotFn): Scalar {
    if (fn === "count") return this.count;
    if (fn === "countNums") return this.nums;
    if (this.error !== null) return this.error;
    switch (fn) {
      case "sum":
        return Math.round(this.sum * 1e10) / 1e10;
      case "average":
        return this.nums ? this.sum / this.nums : Errors.DIV0;
      case "max":
        return this.nums ? this.max : 0;
      case "min":
        return this.nums ? this.min : 0;
      case "product":
        return this.nums ? this.product : 0;
    }
    return null;
  }
}

interface TreeNode {
  key: string;
  value: Scalar;
  children: TreeNode[];
  childMap: Map<string, TreeNode>;
  path: string;
  depth: number;
}

function newNode(key: string, value: Scalar, path: string, depth: number): TreeNode {
  return { key, value, children: [], childMap: new Map(), path, depth };
}

function insertPath(root: TreeNode, keys: string[], values: Scalar[]): void {
  let n = root;
  for (let i = 0; i < keys.length; i++) {
    let c = n.childMap.get(keys[i]);
    if (!c) {
      c = newNode(keys[i], values[i], n.path + "\u0001" + keys[i], i + 1);
      n.childMap.set(keys[i], c);
      n.children.push(c);
    }
    n = c;
  }
}

function sortTree(n: TreeNode, desc: (depth: number) => boolean): void {
  const d = desc(n.depth);
  n.children.sort((a, b) => (d ? -1 : 1) * compareItems(a.value, b.value));
  for (const c of n.children) sortTree(c, desc);
}

/** Filtrelenmiş kayıt dizinleri. */
export function filteredRecords(def: PivotDef, src: PivotSource): number[] {
  const out: number[] = [];
  const pageSel = def.filters.filter((f) => f.selected).map((f) => ({ field: f.field, set: new Set(f.selected!) }));
  const hidden = Object.entries(def.hidden ?? {})
    .filter(([, list]) => list.length)
    .map(([f, list]) => ({ field: Number(f), set: new Set(list) }));
  for (let i = 0; i < src.records.length; i++) {
    const r = src.records[i];
    let ok = true;
    for (const p of pageSel) if (!p.set.has(itemKey(r[p.field] ?? null))) ok = false;
    for (const h of hidden) if (h.set.has(itemKey(r[h.field] ?? null))) ok = false;
    if (ok) out.push(i);
  }
  return out;
}

/* ----------------------------- Yerleşim ----------------------------- */

interface ColEntry {
  /** Sütun ağacı yolu ("" = tümü). */
  path: string;
  vf: number;
  kind: "item" | "subtotal" | "grand";
}

interface RowEntry {
  path: string;
  /** Σ satırlardaysa değer alanı; değilse -1. */
  vf: number;
  kind: "item" | "parent" | "subtotal" | "grand";
  node?: TreeNode;
  depth: number;
}

export function computePivot(def: PivotDef, src: PivotSource): PivotLayout {
  const headers = src.headers;
  const values = def.values;
  const multi = values.length > 1;
  const rowFields = def.rows.filter((f) => f !== VALUES_FIELD);
  const colFields = def.cols.filter((f) => f !== VALUES_FIELD);
  const sigmaRows = multi && !!def.valuesOnRows;
  const sigmaCols = multi && !def.valuesOnRows;
  const cells: PivotCell[] = [];
  const empty = values.length === 0 && rowFields.length === 0 && colFields.length === 0;

  // Filtre alanları (rapor filtresi) tablonun üstünde
  let top = 0;
  for (const f of def.filters) {
    const sel = f.selected;
    let label: Scalar = "(Tümü)";
    if (sel && sel.length === 1) {
      const it = fieldItems(src, f.field).find((x) => x.key === sel[0]);
      label = it ? itemLabel(it.value) : BLANK_ITEM;
    } else if (sel) label = "(Birden Çok Öğe)";
    cells.push({ row: top, col: 0, value: headers[f.field] ?? "", kind: "filterLabel" });
    cells.push({ row: top, col: 1, value: label, kind: "filterValue" });
    top++;
  }
  if (def.filters.length) top++;
  if (empty) return { rows: top, cols: 2, cells, tableRow: top, headerRows: 0, labelCols: 0, empty: true };

  const recs = filteredRecords(def, src);
  // Ağaçlar
  const rowRoot = newNode("", null, "", 0);
  const colRoot = newNode("", null, "", 0);
  const rowKeyOf: string[][] = [];
  const colKeyOf: string[][] = [];
  for (const i of recs) {
    const r = src.records[i];
    const rk = rowFields.map((f) => itemKey(r[f] ?? null));
    const ck = colFields.map((f) => itemKey(r[f] ?? null));
    rowKeyOf[i] = rk;
    colKeyOf[i] = ck;
    insertPath(
      rowRoot,
      rk,
      rowFields.map((f) => r[f] ?? null),
    );
    insertPath(
      colRoot,
      ck,
      colFields.map((f) => r[f] ?? null),
    );
  }
  sortTree(rowRoot, (d) => !!def.sort?.[rowFields[d]]);
  sortTree(colRoot, (d) => !!def.sort?.[colFields[d]]);

  // Toplayıcılar: her satır ön eki × sütun ön eki
  const accs = new Map<string, Acc[]>();
  const accFor = (rp: string, cp: string) => {
    const k = rp + "\u0002" + cp;
    let a = accs.get(k);
    if (!a) accs.set(k, (a = values.map(() => new Acc())));
    return a;
  };
  for (const i of recs) {
    const r = src.records[i];
    const rk = rowKeyOf[i];
    const ck = colKeyOf[i];
    const rps = [""];
    for (let a = 0; a < rk.length; a++) rps.push(rps[a] + "\u0001" + rk[a]);
    const cps = [""];
    for (let b = 0; b < ck.length; b++) cps.push(cps[b] + "\u0001" + ck[b]);
    for (const rp of rps)
      for (const cp of cps) {
        const list = accFor(rp, cp);
        values.forEach((vf, vi) => list[vi].add(r[vf.field] ?? null));
      }
  }
  const valueAt = (rp: string, cp: string, vi: number): Scalar => {
    const a = accs.get(rp + "\u0002" + cp);
    if (!a) return null;
    return a[vi].result(values[vi].fn);
  };
  const fmtOf = (vi: number) => {
    const vf = values[vi];
    if (vf.numFmt) return vf.numFmt;
    if (vf.fn === "count" || vf.fn === "countNums") return undefined;
    return src.numFmts[vf.field];
  };

  const subtotals = def.subtotals;
  const layout = def.layout;
  const compact = layout === "compact";

  // Sütun girdileri
  const colEntries: ColEntry[] = [];
  const valueIdx = values.map((_, i) => i);
  // Değer alanı yoksa sütun öğeleri boş sütunlar olarak gösterilir (vf = -1)
  const colVfs = values.length === 0 ? [-1] : sigmaCols ? valueIdx : [0];
  const walkCols = (n: TreeNode) => {
    if (n.depth === colFields.length) {
      for (const vi of colVfs) colEntries.push({ path: n.path, vf: vi, kind: "item" });
      return;
    }
    for (const c of n.children) walkCols(c);
    // İç içe sütun alanlarında grup alt toplamı (en dış seviyeler)
    if (n.depth > 0 && n.depth < colFields.length && subtotals !== "none") for (const vi of colVfs) colEntries.push({ path: n.path, vf: vi, kind: "subtotal" });
  };
  if (colFields.length) walkCols(colRoot);
  else if (values.length) for (const vi of colVfs) colEntries.push({ path: "", vf: vi, kind: "item" });
  if (colFields.length && def.grandCols && values.length) for (const vi of colVfs) colEntries.push({ path: "", vf: vi, kind: "grand" });

  // Satır girdileri
  const rowEntries: RowEntry[] = [];
  const rowVfs = sigmaRows ? valueIdx : [-1];
  void valueIdx;
  const walkRows = (n: TreeNode) => {
    const isLeafLevel = n.depth === rowFields.length;
    if (n.depth > 0) {
      if (isLeafLevel) {
        if (sigmaRows) {
          rowEntries.push({ path: n.path, vf: -1, kind: "parent", node: n, depth: n.depth });
          for (const vi of rowVfs) rowEntries.push({ path: n.path, vf: vi, kind: "item", node: n, depth: n.depth + 1 });
        } else rowEntries.push({ path: n.path, vf: -1, kind: "item", node: n, depth: n.depth });
        return;
      }
      // Grup satırı (alt toplam üstte ise değer taşır); tablo biçiminde grup etiketi ilk alt satırdadır
      const topTotals = subtotals === "top" && layout !== "tabular" && !sigmaRows;
      if (layout !== "tabular") rowEntries.push({ path: n.path, vf: -1, kind: topTotals ? "subtotal" : "parent", node: n, depth: n.depth });
    }
    for (const c of n.children) walkRows(c);
    if (n.depth > 0 && !isLeafLevel && (subtotals === "bottom" || (subtotals === "top" && (layout === "tabular" || sigmaRows)))) {
      for (const vi of rowVfs) rowEntries.push({ path: n.path, vf: vi, kind: "subtotal", node: n, depth: n.depth });
    }
    if (n.depth > 0 && def.blankRows && n.depth === 1 && rowFields.length > 1) rowEntries.push({ path: "\u0003blank", vf: -1, kind: "parent", depth: 0 });
  };
  if (rowFields.length) walkRows(rowRoot);
  else for (const vi of rowVfs) rowEntries.push({ path: "", vf: vi, kind: "item", depth: 0 });
  if (rowFields.length && def.grandRows) for (const vi of rowVfs) rowEntries.push({ path: "", vf: vi, kind: "grand", depth: 0 });

  // Etiket sütunları
  const labelCols = rowFields.length === 0 ? (colFields.length || sigmaRows ? 1 : 0) : compact ? 1 : rowFields.length + (sigmaRows ? 1 : 0);
  const headerRows = colFields.length ? colFields.length + 1 + (sigmaCols ? 1 : 0) : 1;
  const T = top;
  const captionSingle = values.length === 1 ? valueCaption(values[0], headers) : values.length === 0 ? "" : "Değerler";

  // Başlık alanı
  if (colFields.length) {
    if (labelCols > 0) cells.push({ row: T, col: 0, value: values.length === 1 ? captionSingle : null, kind: "caption" });
    cells.push({ row: T, col: labelCols, value: def.colHeaders ? (compact ? "Sütun Etiketleri" : headers[colFields[0]]) : null, kind: "colLabelsHeader" });
    // Sütun öğe başlıkları
    const levels = colFields.length + (sigmaCols ? 1 : 0);
    let prevPath: string[] = [];
    colEntries.forEach((ce, i) => {
      const c = labelCols + i;
      const parts = ce.path.split("\u0001").slice(1);
      const node = (p: number) => {
        let n = colRoot;
        for (let k = 0; k <= p; k++) n = n.childMap.get(parts[k])!;
        return n;
      };
      if (ce.kind === "grand") {
        cells.push({ row: T + 1, col: c, value: multi && ce.vf >= 0 ? `Toplam ${valueCaption(values[ce.vf], headers)}` : "Genel Toplam", kind: "grandColHeader" });
        prevPath = [];
        return;
      }
      if (ce.kind === "subtotal") {
        const n = node(parts.length - 1);
        cells.push({ row: T + parts.length, col: c, value: `${labelText(n.value)} Toplam${multi && sigmaCols && ce.vf >= 0 ? " " + valueCaption(values[ce.vf], headers) : ""}`, kind: "grandColHeader" });
        prevPath = [];
        return;
      }
      const prefix = (p: string[], lv: number) => p.slice(0, lv + 1).join("\u0001");
      for (let lv = 0; lv < parts.length; lv++) {
        // Grup etiketi yalnız grubun ilk sütununda
        if (prevPath.length > lv && prefix(prevPath, lv) === prefix(parts, lv)) continue;
        cells.push({ row: T + 1 + lv, col: c, value: itemLabel(node(lv).value), kind: "colItem" });
      }
      if (sigmaCols && ce.vf >= 0) cells.push({ row: T + levels, col: c, value: valueCaption(values[ce.vf], headers), kind: "colItem" });
      prevPath = parts;
    });
    // Satır etiketleri başlığı (son başlık satırında)
    const hr = T + headerRows - 1;
    if (labelCols > 0) {
      if (compact) {
        if (rowFields.length || sigmaRows) cells.push({ row: hr, col: 0, value: def.rowHeaders ? "Satır Etiketleri" : null, kind: "rowLabelsHeader" });
      } else rowFields.forEach((f, i) => cells.push({ row: hr, col: i, value: headers[f], kind: "fieldHeader" }));
    }
  } else {
    // Sütun alanı yok: tek başlık satırı
    if (labelCols > 0) {
      if (compact || rowFields.length === 0) cells.push({ row: T, col: 0, value: rowFields.length || sigmaRows ? (def.rowHeaders ? "Satır Etiketleri" : null) : null, kind: "rowLabelsHeader" });
      else rowFields.forEach((f, i) => cells.push({ row: T, col: i, value: headers[f], kind: "fieldHeader" }));
    }
    colEntries.forEach((ce, i) => {
      if (ce.vf >= 0) cells.push({ row: T, col: labelCols + i, value: sigmaRows ? "Değerler" : valueCaption(values[ce.vf], headers), kind: "colItem" });
    });
  }

  // Gövde
  const body = T + headerRows;
  let band = 0;
  let lastLabels: Scalar[] = [];
  rowEntries.forEach((re, i) => {
    const r = body + i;
    if (re.path === "\u0003blank") return;
    const parts = re.path.split("\u0001").slice(1);
    const bold = re.kind === "subtotal" || re.kind === "grand" || (re.kind === "parent" && !!re.node && re.node.depth < rowFields.length);
    // Etiket
    if (re.kind === "grand") {
      cells.push({ row: r, col: 0, value: sigmaRows ? `Toplam ${valueCaption(values[re.vf], headers)}` : "Genel Toplam", kind: "grandLabel", bold: true });
    } else if (re.kind === "subtotal" && (subtotals === "bottom" || layout === "tabular" || sigmaRows) && re.node) {
      const col = compact ? 0 : re.node.depth - 1;
      const label = `${labelText(re.node.value)} Toplam${sigmaRows && re.vf >= 0 ? " " + valueCaption(values[re.vf], headers) : ""}`;
      cells.push({ row: r, col, value: label, kind: "subtotalLabel", indent: compact ? re.node.depth - 1 : 0, bold: true });
    } else if (re.vf >= 0 && sigmaRows && re.kind === "item") {
      cells.push({ row: r, col: compact ? 0 : rowFields.length, value: valueCaption(values[re.vf], headers), kind: "rowItem", indent: compact ? re.depth - 1 : 0 });
    } else if (re.node) {
      const n = re.node;
      if (compact) cells.push({ row: r, col: 0, value: itemLabel(n.value), kind: re.kind === "item" ? "rowItem" : "rowParent", indent: n.depth - 1, bold });
      else {
        // Tablo/anahat: her alan kendi sütununda; üst öğeler grubun ilk satırında (ya da yinelenir)
        const labels: Scalar[] = [];
        let m = rowRoot;
        for (let k = 0; k < parts.length; k++) {
          m = m.childMap.get(parts[k])!;
          labels.push(itemLabel(m.value));
        }
        const keyTo = (list: Scalar[], k: number) => list.slice(0, k + 1).map(String).join("\u0001");
        for (let k = 0; k < labels.length; k++) {
          const isOwn = k === labels.length - 1;
          if (!isOwn && keyTo(lastLabels, k) === keyTo(labels, k) && !def.repeatLabels) continue;
          if (layout === "outline" && !isOwn && re.kind === "item" && !def.repeatLabels) continue;
          cells.push({ row: r, col: k, value: labels[k], kind: isOwn && re.kind === "item" ? "rowItem" : "rowParent", bold: bold || !isOwn });
        }
        lastLabels = labels;
      }
    }
    if (re.kind === "parent") return;
    // Değerler
    const isTotalRow = re.kind === "subtotal" || re.kind === "grand";
    if (!isTotalRow) band++;
    colEntries.forEach((ce, j) => {
      const vi = sigmaRows ? re.vf : ce.vf;
      if (vi < 0 || vi >= values.length) return;
      const v = valueAt(re.path, ce.path, vi);
      const grand = re.kind === "grand" || ce.kind === "grand";
      cells.push({ row: r, col: labelCols + j, value: v, kind: grand ? "grandValue" : isTotalRow || ce.kind === "subtotal" ? "subtotalValue" : "value", numFmt: fmtOf(vi), bold: isTotalRow, band });
    });
  });

  const rows = Math.max(top + 1, body + rowEntries.length);
  const cols = Math.max(2, labelCols + colEntries.length);
  return { rows, cols, cells, tableRow: T, headerRows, labelCols, empty: false };
}

function labelText(v: Scalar): string {
  const l = itemLabel(v);
  return typeof l === "number" ? String(l).replace(".", ",") : String(l);
}

/** Varsayılan yeni PivotTable tanımı. */
export function newPivotDef(id: string, name: string, sourceSheet: number, source: PivotDef["source"], anchor: PivotDef["anchor"]): PivotDef {
  return {
    id,
    name,
    sourceSheet,
    source,
    anchor,
    rows: [],
    cols: [],
    values: [],
    filters: [],
    subtotals: "top",
    grandRows: true,
    grandCols: true,
    layout: "compact",
    style: "PivotStyleLight16",
    rowHeaders: true,
    colHeaders: true,
    bandedRows: false,
    bandedCols: false,
  };
}

/** Alanın sayısal olup olmadığı (Değerler'e mi Satırlar'a mı ekleneceği). */
export function isNumericField(src: PivotSource, field: number): boolean {
  let n = 0;
  let t = 0;
  for (const r of src.records.slice(0, 500)) {
    const v = r[field];
    if (typeof v === "number") n++;
    else if (v !== null && v !== "") t++;
  }
  return n > 0 && n >= t;
}
