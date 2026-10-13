// Tablo tutamaçları (Word'deki gibi): sol üstte taşıma tutamacı (✥; tıklayınca tabloyu seçer, sürükleyince
// tabloyu taşır) ve sağ altta boyutlandırma tutamacı (sütun genişliklerini orantılı ölçekler).
// Ortalı ya da sağa yaslı tablo sağ kenarından boyutlandırılınca sol kenarı yerinde kalır; bırakınca
// tablo sola yaslı olur ve ölçülen sol konum girinti olarak yazılır (tek geri alma adımı).
import { Extension } from "@tiptap/core";
import { closeHistory } from "@tiptap/pm/history";
import { Plugin, PluginKey, TextSelection, type Transaction } from "@tiptap/pm/state";
import { CellSelection, TableMap, columnResizingPluginKey } from "@tiptap/pm/tables";
import type { EditorView } from "@tiptap/pm/view";
import { anchorTableLeft } from "./table";

/** Tabloyu (tüm hücreleriyle) seçer. */
export function selectTable(view: EditorView, tablePos: number) {
  const table = view.state.doc.nodeAt(tablePos);
  if (!table || table.type.name !== "table") return;
  const map = TableMap.get(table);
  const first = tablePos + 1 + map.map[0];
  const last = tablePos + 1 + map.map[map.map.length - 1];
  view.dispatch(view.state.tr.setSelection(CellSelection.create(view.state.doc, first, last)));
}

/** Tabloyu belgede başka bir üst düzey blok konumuna taşır. */
export function moveTable(view: EditorView, tablePos: number, target: number): boolean {
  const { state } = view;
  const table = state.doc.nodeAt(tablePos);
  if (!table) return false;
  const $t = state.doc.resolve(Math.max(0, Math.min(target, state.doc.content.size)));
  if ($t.depth < 1) return false;
  // Hedef: imlecin bulunduğu üst düzey bloğun önü (belge > bölüm > blok).
  const depth = Math.min(2, $t.depth);
  let at = depth >= 2 ? $t.before(2) : $t.pos;
  if (at > tablePos && at < tablePos + table.nodeSize) return false;
  const tr = state.tr.delete(tablePos, tablePos + table.nodeSize);
  at = tr.mapping.map(at, -1);
  tr.insert(at, table);
  tr.setSelection(TextSelection.near(tr.doc.resolve(at + 3)));
  view.dispatch(tr.scrollIntoView());
  return true;
}

/** Tablonun yakınlaştırma ölçeği (ekran pikseli / CSS pikseli). */
function zoomOf(el: HTMLElement): number {
  const r = el.getBoundingClientRect().width;
  return el.offsetWidth ? r / el.offsetWidth : 1;
}

/**
 * Sütun genişlikleri (CSS px). Hücrenin colwidth özniteliği varsa o; yoksa ekrandaki
 * tek sütunluk hücrenin genişliği (yakınlaştırma düzeltilmiş). Birleştirilmiş hücreler
 * yalnız boş kalan sütunlara bölüştürülür.
 */
export function measureColumns(view: EditorView, tablePos: number): number[] {
  const table = view.state.doc.nodeAt(tablePos);
  if (!table || table.type.name !== "table") return [];
  const map = TableMap.get(table);
  const widths: (number | null)[] = Array.from({ length: map.width }, () => null);
  const cells = Array.from(new Set(map.map));
  for (const off of cells) {
    const cw = table.nodeAt(off)?.attrs.colwidth as number[] | null | undefined;
    const rect = map.findCell(off);
    for (let c = rect.left; c < rect.right; c++) if (cw?.[c - rect.left] && widths[c] == null) widths[c] = cw[c - rect.left];
  }
  const dom = view.nodeDOM(tablePos) as HTMLElement | null;
  const zoom = dom ? zoomOf(dom) || 1 : 1;
  const cellWidth = (off: number) => (view.nodeDOM(tablePos + 1 + off) as HTMLElement | null)?.getBoundingClientRect().width ?? 0;
  // Önce tek sütunlu hücreler; birleştirilmiş hücreler ikinci turda kalanı bölüştürür.
  for (const off of cells) {
    const rect = map.findCell(off);
    if (rect.right - rect.left !== 1 || widths[rect.left] != null) continue;
    const w = cellWidth(off) / zoom;
    if (w) widths[rect.left] = w;
  }
  for (const off of cells) {
    const rect = map.findCell(off);
    const missing = Array.from({ length: rect.right - rect.left }, (_, i) => rect.left + i).filter((c) => widths[c] == null);
    if (!missing.length) continue;
    const known = Array.from({ length: rect.right - rect.left }, (_, i) => widths[rect.left + i] ?? 0).reduce((x, y) => x + y, 0);
    const w = cellWidth(off) / zoom;
    for (const c of missing) widths[c] = Math.max(20, (w - known) / missing.length);
  }
  return widths.map((w) => Math.max(20, Math.round(w ?? 100)));
}

/** Her hücrenin colwidth özniteliğini verilen sütun genişliklerine göre yazar. */
export function columnWidthsTr(view: EditorView, tablePos: number, widths: number[]) {
  const table = view.state.doc.nodeAt(tablePos);
  const tr = view.state.tr;
  if (!table) return tr;
  const map = TableMap.get(table);
  for (const off of new Set(map.map)) {
    const cell = table.nodeAt(off);
    if (!cell) continue;
    const rect = map.findCell(off);
    const colwidth = widths.slice(rect.left, rect.right).map((w) => Math.max(20, Math.round(w)));
    const cur = cell.attrs.colwidth as number[] | null;
    if (cur && cur.length === colwidth.length && cur.every((w, i) => w === colwidth[i])) continue;
    tr.setNodeMarkup(tablePos + 1 + off, undefined, { ...cell.attrs, colwidth });
  }
  return tr;
}

/**
 * Boyutlandırmadan önce genişliği tanımsız sütunları ekrandaki genişlikleriyle sabitler
 * (geri al geçmişine girmez). Esnek sütunlar sürüklerken kendiliğinden yeniden dağıldığı
 * için tablo zıplıyordu; sabit sütunlarla yalnız sürüklenen kenar hareket eder.
 */
export function freezeColumns(view: EditorView, tablePos: number): number[] {
  const widths = measureColumns(view, tablePos);
  const tr = columnWidthsTr(view, tablePos, widths);
  if (tr.docChanged) view.dispatch(tr.setMeta("addToHistory", false));
  return widths;
}

/**
 * Ortalı/sağa yaslı tablonun sarmalayıcıya göre sol konumu (CSS px; yakınlaştırma düzeltilmiş).
 * Sola yaslı tabloda null: sol kenar zaten yerinde kalır.
 */
export function floatingTableLeft(view: EditorView, tablePos: number, table: HTMLElement): number | null {
  const align = view.state.doc.nodeAt(tablePos)?.attrs.align;
  const wrap = table.parentElement;
  if ((align !== "center" && align !== "right") || !wrap) return null;
  return (table.getBoundingClientRect().left - wrap.getBoundingClientRect().left) / (zoomOf(table) || 1);
}

/**
 * Boyutlandırma işlemine tablonun yeni yerleşimini ekler: sabitlenen sol kenar girinti olur, tablo
 * sola yaslanır. Sütun genişlikleriyle aynı adımda yazıldığından Geri Al ikisini birlikte geri alır.
 */
export function settleTableLeft(tr: Transaction, tablePos: number, left: number | null): Transaction {
  const node = tr.doc.nodeAt(tablePos);
  if (left === null || !node || node.type.name !== "table") return tr;
  return tr.setNodeMarkup(tablePos, undefined, { ...node.attrs, align: null, indent: Math.round(left) });
}

/** Sütun genişliklerini orantılı ölçekler. */
export function scaleTable(view: EditorView, tablePos: number, factor: number) {
  if (!Number.isFinite(factor) || factor <= 0) return;
  const widths = measureColumns(view, tablePos);
  if (!widths.length) return;
  view.dispatch(closeHistory(columnWidthsTr(view, tablePos, widths.map((w) => w * factor))));
}

class HandlesView {
  private move: HTMLDivElement;
  private resize: HTMLDivElement;
  private tablePos = -1;
  private tableEl: HTMLElement | null = null;
  private hideTimer = 0;

  constructor(private view: EditorView) {
    this.move = document.createElement("div");
    this.move.className = "bl-tbl-move";
    this.move.title = "Tabloyu seçmek için tıklayın, taşımak için sürükleyin";
    this.move.textContent = "✥";
    this.resize = document.createElement("div");
    this.resize.className = "bl-tbl-resize";
    this.resize.title = "Tabloyu yeniden boyutlandırın";
    for (const el of [this.move, this.resize]) {
      el.setAttribute("aria-hidden", "true");
      el.style.display = "none";
      document.body.appendChild(el);
      el.addEventListener("mouseenter", () => window.clearTimeout(this.hideTimer));
    }
    view.dom.addEventListener("mousemove", this.onMove);
    view.dom.addEventListener("mouseleave", this.onLeave);
    // Yakalama evresinde: tablo eklentisinin sürükleme başlangıcından önce çalışır.
    view.dom.addEventListener("mousedown", this.onBorderDown, true);
    document.addEventListener("scroll", this.reposition, true);
    this.move.addEventListener("mousedown", this.onMoveDown);
    this.resize.addEventListener("mousedown", this.onResizeDown);
  }

  private onMove = (e: MouseEvent) => {
    const t = (e.target as HTMLElement).closest?.("table");
    if (!t || !this.view.dom.contains(t)) return;
    window.clearTimeout(this.hideTimer);
    const cur = this.tablePos >= 0 ? (this.view.nodeDOM(this.tablePos) as HTMLElement | null) : null;
    let tablePos = cur && cur.contains(t) ? this.tablePos : -1;
    if (tablePos < 0) {
      this.view.state.doc.descendants((n, p) => {
        if (tablePos >= 0) return false;
        if (n.type.name !== "table") return true;
        const d = this.view.nodeDOM(p) as HTMLElement | null;
        if (d && d.contains(t)) tablePos = p;
        return false;
      });
    }
    if (tablePos < 0) return;
    this.tablePos = tablePos;
    this.tableEl = t;
    this.reposition();
  };

  private onLeave = (e: MouseEvent) => {
    const to = e.relatedTarget as HTMLElement | null;
    if (to && (to === this.move || to === this.resize)) return;
    this.hideTimer = window.setTimeout(() => this.hide(), 600);
  };

  private hide() {
    this.move.style.display = "none";
    this.resize.style.display = "none";
  }

  reposition = () => {
    const t = this.tableEl;
    if (!t || !t.isConnected || !this.view.editable) return this.hide();
    const r = t.getBoundingClientRect();
    if (r.width === 0) return this.hide();
    Object.assign(this.move.style, { display: "flex", left: `${r.left - 16}px`, top: `${r.top - 16}px` });
    Object.assign(this.resize.style, { display: "block", left: `${r.right + 1}px`, top: `${r.bottom + 1}px` });
  };

  private onMoveDown = (e: MouseEvent) => {
    e.preventDefault();
    const pos = this.tablePos;
    if (pos < 0) return;
    selectTable(this.view, pos);
    this.view.focus();
    const sx = e.clientX;
    const sy = e.clientY;
    let moved = false;
    const line = document.createElement("div");
    line.className = "bl-tbl-drop";
    const onMove = (ev: MouseEvent) => {
      if (!moved && Math.hypot(ev.clientX - sx, ev.clientY - sy) < 6) return;
      moved = true;
      const hit = this.view.posAtCoords({ left: ev.clientX, top: ev.clientY });
      if (!hit) return;
      const c = this.view.coordsAtPos(hit.pos);
      const pm = this.view.dom.getBoundingClientRect();
      Object.assign(line.style, { left: `${pm.left}px`, width: `${pm.width}px`, top: `${c.top - 2}px` });
      if (!line.isConnected) document.body.appendChild(line);
    };
    const onUp = (ev: MouseEvent) => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      line.remove();
      if (!moved) return;
      const hit = this.view.posAtCoords({ left: ev.clientX, top: ev.clientY });
      if (hit) moveTable(this.view, pos, hit.pos);
      this.hide();
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  };

  private onResizeDown = (e: MouseEvent) => {
    e.preventDefault();
    const pos = this.tablePos;
    if (!this.tableEl || pos < 0) return;
    // Sürükleme başlayınca sütunlar sabitlenir; böylece yalnız orantılı ölçek değişir.
    const start = measureColumns(this.view, pos);
    let frozen = false;
    const t = (this.view.nodeDOM(pos) as HTMLElement | null)?.querySelector("table") ?? null;
    if (!t || !start.length) return;
    this.tableEl = t;
    const cols = Array.from(t.querySelectorAll<HTMLElement>(":scope > colgroup > col"));
    const total = start.reduce((x, y) => x + y, 0);
    const zoom = zoomOf(t) || 1;
    // Ortalı tablo: sol kenar şimdiki yerinde kalsın (sürüklerken yeniden ortalanmasın).
    const left = floatingTableLeft(this.view, pos, t);
    const sx = e.clientX;
    const factorAt = (x: number) => Math.max(60, total + (x - sx) / zoom) / total;
    const show = (f: number) => {
      start.forEach((w, i) => {
        if (cols[i]) cols[i].style.width = `${Math.max(20, Math.round(w * f))}px`;
      });
      t.style.width = `${cols.length ? start.reduce((sum, w) => sum + Math.max(20, Math.round(w * f)), 0) : Math.round(total * f)}px`;
      t.style.minWidth = "";
      this.reposition();
    };
    const onMove = (ev: MouseEvent) => {
      if (!frozen) {
        if (Math.abs(ev.clientX - sx) < 2) return;
        frozen = true;
        if (left !== null) anchorTableLeft(t, left);
        freezeColumns(this.view, pos);
      }
      show(factorAt(ev.clientX));
    };
    const onUp = (ev: MouseEvent) => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      if (!frozen) return;
      const f = factorAt(ev.clientX);
      if (left !== null) anchorTableLeft(t, null);
      if (Math.abs(f - 1) > 0.005) this.view.dispatch(closeHistory(settleTableLeft(columnWidthsTr(this.view, pos, start.map((w) => w * f)), pos, left)));
      else show(1);
      this.reposition();
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  };

  /**
   * Sütun kenarı sürükleme (Word'deki gibi): kenarın iki yanındaki sütunlar birlikte değişir,
   * tablo genişliği korunur; son sütunun sağ kenarı tabloyu genişletir. Tablo eklentisinin
   * kendi sürüklemesi esnek sütunları en küçük genişliğe çöktürdüğü için olay burada alınır.
   */
  private onBorderDown = (e: MouseEvent) => {
    if (e.button !== 0 || !this.view.editable || !this.view.dom.classList.contains("resize-cursor")) return;
    const handle = columnResizingPluginKey.getState(this.view.state)?.activeHandle ?? -1;
    if (handle < 0) return;
    const $cell = this.view.state.doc.resolve(handle);
    const table = $cell.node(-1);
    if (!table || table.type.name !== "table") return;
    const tablePos = $cell.before(-1);
    const col = TableMap.get(table).findCell(handle - $cell.start(-1)).right - 1;
    e.preventDefault();
    e.stopPropagation();
    const start = measureColumns(this.view, tablePos);
    let frozen = false;
    const t = (this.view.nodeDOM(tablePos) as HTMLElement | null)?.querySelector("table");
    if (!t || col < 0 || col >= start.length) return;
    const cols = Array.from(t.querySelectorAll<HTMLElement>(":scope > colgroup > col"));
    const zoom = zoomOf(t) || 1;
    // Son sütunun sağ kenarı tablo genişliğini değiştirir: ortalı tablonun sol kenarı yerinde kalsın.
    const left = col === start.length - 1 ? floatingTableLeft(this.view, tablePos, t) : null;
    const sx = e.clientX;
    let widths = start;
    const widthsAt = (x: number) => {
      const dx = (x - sx) / zoom;
      const next = start.slice();
      if (col < start.length - 1) {
        const pair = start[col] + start[col + 1];
        next[col] = Math.max(20, Math.min(pair - 20, start[col] + dx));
        next[col + 1] = pair - next[col];
      } else next[col] = Math.max(20, start[col] + dx);
      return next.map((w) => Math.round(w));
    };
    const show = (ws: number[]) => {
      ws.forEach((w, i) => {
        if (cols[i]) cols[i].style.width = `${w}px`;
      });
      t.style.width = `${ws.reduce((x, y) => x + y, 0)}px`;
      t.style.minWidth = "";
      this.reposition();
    };
    document.body.classList.add("bl-col-resizing");
    const onMove = (ev: MouseEvent) => {
      // Sütunlar ancak gerçekten sürüklenince sabitlenir; tıklayıp bırakmak belgeyi değiştirmez.
      if (!frozen) {
        if (Math.abs(ev.clientX - sx) < 2) return;
        frozen = true;
        if (left !== null) anchorTableLeft(t, left);
        freezeColumns(this.view, tablePos);
      }
      widths = widthsAt(ev.clientX);
      show(widths);
    };
    const onUp = (ev: MouseEvent) => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      document.body.classList.remove("bl-col-resizing");
      if (!frozen) return;
      widths = widthsAt(ev.clientX);
      if (left !== null) anchorTableLeft(t, null);
      if (widths.some((w, i) => w !== start[i])) this.view.dispatch(closeHistory(settleTableLeft(columnWidthsTr(this.view, tablePos, widths), tablePos, left)));
      else show(start);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  };

  update() {
    if (this.move.style.display !== "none") this.reposition();
  }

  destroy() {
    this.view.dom.removeEventListener("mousemove", this.onMove);
    this.view.dom.removeEventListener("mouseleave", this.onLeave);
    this.view.dom.removeEventListener("mousedown", this.onBorderDown, true);
    document.removeEventListener("scroll", this.reposition, true);
    this.move.remove();
    this.resize.remove();
    window.clearTimeout(this.hideTimer);
  }
}

export const TableHandles = Extension.create({
  name: "belgeTableHandles",
  addProseMirrorPlugins() {
    return [new Plugin({ key: new PluginKey("belgeTableHandles"), view: (view) => new HandlesView(view) })];
  },
});
