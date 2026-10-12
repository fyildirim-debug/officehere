// Tablo tutamaçları (Word'deki gibi): sol üstte taşıma tutamacı (✥; tıklayınca tabloyu seçer, sürükleyince
// tabloyu taşır) ve sağ altta boyutlandırma tutamacı (sütun genişliklerini orantılı ölçekler).
import { Extension } from "@tiptap/core";
import type { Node as PMNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, TextSelection } from "@tiptap/pm/state";
import { CellSelection, TableMap } from "@tiptap/pm/tables";
import type { EditorView } from "@tiptap/pm/view";

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

/** Sütun genişliklerini orantılı ölçekler. */
export function scaleTable(view: EditorView, tablePos: number, factor: number) {
  const table = view.state.doc.nodeAt(tablePos);
  if (!table || !Number.isFinite(factor) || factor <= 0) return;
  const tr = view.state.tr;
  const domTable = (view.nodeDOM(tablePos) as HTMLElement | null)?.querySelector("table");
  const widths = Array.from(domTable?.querySelectorAll("tr:first-child > *") ?? []).map((c) => (c as HTMLElement).getBoundingClientRect().width);
  table.descendants((cell: PMNode, off: number) => {
    if (cell.type.name !== "tableCell" && cell.type.name !== "tableHeader") return true;
    const cw = (cell.attrs.colwidth as number[] | null) ?? null;
    const span = (cell.attrs.colspan as number) || 1;
    const base = cw && cw.every((w) => w > 0) ? cw : Array.from({ length: span }, () => (widths[0] ?? 100) / Math.max(1, span));
    tr.setNodeMarkup(tablePos + 1 + off, undefined, { ...cell.attrs, colwidth: base.map((w) => Math.max(20, Math.round(w * factor))) });
    return false;
  });
  view.dispatch(tr);
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
    const t = this.tableEl;
    const pos = this.tablePos;
    if (!t || pos < 0) return;
    const startW = t.getBoundingClientRect().width;
    const sx = e.clientX;
    const scale = startW / (t.offsetWidth || startW);
    const onMove = (ev: MouseEvent) => {
      t.style.width = `${Math.max(60, (startW + ev.clientX - sx) / scale)}px`;
      this.reposition();
    };
    const onUp = (ev: MouseEvent) => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      const factor = Math.max(60, startW + ev.clientX - sx) / startW;
      if (Math.abs(factor - 1) > 0.01) scaleTable(this.view, pos, factor);
      else t.style.width = "";
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
