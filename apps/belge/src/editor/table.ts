// Tablo: stil (Tablo Tasarımı galerisi), stil seçenekleri, hücre gölgelendirmesi ve dikey hizalama.
import { Table, TableCell, TableHeader, TableRow } from "@tiptap/extension-table";
import type { Node as PMNode } from "@tiptap/pm/model";
import type { NodeView, ViewMutationRecord } from "@tiptap/pm/view";
import { cellPadCss, type CellMargins } from "./tableTools";

const cellMarginsAttr = (m: CellMargins) => [m.top, m.right, m.bottom, m.left].join(",");

function parseCellMargins(v: string | null): CellMargins | null {
  const n = (v ?? "").split(",").map(Number);
  return n.length === 4 && n.every((x) => Number.isFinite(x)) ? { top: n[0], right: n[1], bottom: n[2], left: n[3] } : null;
}

export interface TableStyleDef {
  id: string;
  label: string;
  /** Başlık satırı dolgusu, yazı rengi, bant rengi, kenarlık rengi. */
  header?: string;
  headerText?: string;
  band?: string;
  border: string;
  outerOnly?: boolean;
  noBorders?: boolean;
}

export const TABLE_STYLES: TableStyleDef[] = [
  { id: "grid", label: "Tablo Kılavuzu", border: "#000000" },
  { id: "plain", label: "Düz Tablo", border: "#BFBFBF" },
  { id: "none", label: "Kenarlıksız", border: "transparent", noBorders: true },
  { id: "grid-blue", label: "Kılavuz Tablo 4 - Vurgu 1", header: "#156082", headerText: "#FFFFFF", band: "#C1E4F5", border: "#45B0E1" },
  { id: "grid-orange", label: "Kılavuz Tablo 4 - Vurgu 2", header: "#E97132", headerText: "#FFFFFF", band: "#FBE2D5", border: "#F1A983" },
  { id: "grid-green", label: "Kılavuz Tablo 4 - Vurgu 3", header: "#196B24", headerText: "#FFFFFF", band: "#C1F0C7", border: "#47D459" },
  { id: "grid-gray", label: "Kılavuz Tablo 4", header: "#000000", headerText: "#FFFFFF", band: "#CCCCCC", border: "#666666" },
  { id: "list-blue", label: "Liste Tablo 3 - Vurgu 1", header: "#156082", headerText: "#FFFFFF", border: "#156082", outerOnly: true },
  { id: "light-blue", label: "Açık Liste - Vurgu 1", header: "#DAE9F7", headerText: "#0F4761", band: "#F2F7FC", border: "#94C3E8" },
  { id: "light-purple", label: "Kılavuz Tablo 2 - Vurgu 5", header: "#F2CEEF", headerText: "#000000", band: "#FAEBF8", border: "#D86DCB" },
];

/** Tablo düğümü için kendi görünümümüz: stil özniteliklerini <table> öğesine yansıtır. */
class BelgeTableView implements NodeView {
  dom: HTMLElement;
  table: HTMLTableElement;
  colgroup: HTMLTableColElement;
  contentDOM: HTMLElement;

  constructor(
    public node: PMNode,
    public cellMinWidth: number,
  ) {
    this.dom = document.createElement("div");
    this.dom.className = "tableWrapper";
    this.table = this.dom.appendChild(document.createElement("table"));
    this.colgroup = this.table.appendChild(document.createElement("colgroup"));
    this.contentDOM = this.table.appendChild(document.createElement("tbody"));
    this.apply();
  }

  private apply() {
    const a = this.node.attrs;
    this.table.className = "bl-table";
    this.table.setAttribute("data-table-style", (a.tableStyle as string) ?? "grid");
    this.table.setAttribute("data-header-row", String(!!a.headerRow));
    this.table.setAttribute("data-banded-rows", String(!!a.bandedRows));
    this.table.setAttribute("data-first-col", String(!!a.firstColumn));
    this.table.setAttribute("data-total-row", String(!!a.totalRow));
    const pad = a.cellMargins as CellMargins | null;
    if (pad) {
      this.table.setAttribute("data-cpad", cellMarginsAttr(pad));
      this.table.style.setProperty("--bl-cpad", cellPadCss(pad));
    } else {
      this.table.removeAttribute("data-cpad");
      this.table.style.removeProperty("--bl-cpad");
    }
    if (a.align) this.table.setAttribute("data-align", a.align as string);
    else this.table.removeAttribute("data-align");
    this.dom.style.justifyContent = a.align === "center" ? "center" : a.align === "right" ? "flex-end" : "";
    this.applyIndent();
    updateColumns(this.node, this.colgroup, this.table, this.cellMinWidth);
  }

  /**
   * Sol girinti (yalnız sola yaslı tabloda). Boyutlandırma sürerken (data-bl-anchor) tablo sol kenarı
   * ölçülen yerde tutulur: ortalı/sağa yaslı tablo sağ kenarından küçültülünce yeniden ortalanmaz.
   */
  private applyIndent() {
    const anchor = this.table.dataset.blAnchor;
    const indent = Number(this.node.attrs.indent) || 0;
    if (indent) this.table.setAttribute("data-indent", String(indent));
    else this.table.removeAttribute("data-indent");
    if (anchor !== undefined) {
      this.table.style.marginLeft = `${anchor}px`;
      this.table.style.marginRight = "auto";
    } else {
      this.table.style.marginLeft = (!this.node.attrs.align || this.node.attrs.align === "left") && indent ? `${indent}px` : "";
      this.table.style.marginRight = "";
    }
  }

  update(node: PMNode) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.apply();
    return true;
  }

  ignoreMutation(m: ViewMutationRecord) {
    return m.type === "attributes" && (m.target === this.table || this.colgroup.contains(m.target as globalThis.Node));
  }
}

/**
 * Boyutlandırma sürerken tablonun sol kenarını sarmalayıcıya göre `px` uzaklıkta tutar (null: bırakır).
 * Yalnız <table> öğesinin öznitelikleri değişir (düğüm görünümü bu değişiklikleri yok sayar).
 */
export function anchorTableLeft(table: HTMLElement, px: number | null) {
  if (px === null) {
    delete table.dataset.blAnchor;
    table.style.marginLeft = "";
    table.style.marginRight = "";
    return;
  }
  table.dataset.blAnchor = String(Math.round(px * 100) / 100);
  table.style.marginLeft = `${table.dataset.blAnchor}px`;
  table.style.marginRight = "auto";
}

export function updateColumns(node: PMNode, colgroup: HTMLElement, table: HTMLTableElement, cellMinWidth: number, overrideCol?: number, overrideValue?: number) {
  let total = 0;
  let fixed = true;
  let next = colgroup.firstChild as HTMLElement | null;
  const row = node.firstChild;
  if (row) {
    for (let i = 0, col = 0; i < row.childCount; i++) {
      const { colspan, colwidth } = row.child(i).attrs as { colspan: number; colwidth: number[] | null };
      for (let j = 0; j < colspan; j++, col++) {
        const w = overrideCol === col ? overrideValue : colwidth?.[j];
        const css = w ? `${w}px` : "";
        total += w || cellMinWidth;
        if (!w) fixed = false;
        if (!next) {
          const c = document.createElement("col");
          c.style.width = css;
          colgroup.appendChild(c);
        } else {
          if (next.style.width !== css) next.style.width = css;
          next = next.nextSibling as HTMLElement | null;
        }
      }
    }
  }
  while (next) {
    const after = next.nextSibling as HTMLElement | null;
    next.remove();
    next = after;
  }
  if (fixed) {
    table.style.width = `${total}px`;
    table.style.minWidth = "";
  } else {
    table.style.width = "100%";
    table.style.minWidth = `${total}px`;
  }
}

export const BelgeTable = Table.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      tableStyle: { default: "grid", parseHTML: (el) => el.getAttribute("data-table-style") ?? "grid", renderHTML: (a) => ({ "data-table-style": a.tableStyle }) },
      headerRow: { default: true, parseHTML: (el) => el.getAttribute("data-header-row") !== "false", renderHTML: (a) => ({ "data-header-row": String(a.headerRow) }) },
      bandedRows: { default: true, parseHTML: (el) => el.getAttribute("data-banded-rows") !== "false", renderHTML: (a) => ({ "data-banded-rows": String(a.bandedRows) }) },
      firstColumn: { default: false, parseHTML: (el) => el.getAttribute("data-first-col") === "true", renderHTML: (a) => ({ "data-first-col": String(a.firstColumn) }) },
      totalRow: { default: false, parseHTML: (el) => el.getAttribute("data-total-row") === "true", renderHTML: (a) => ({ "data-total-row": String(a.totalRow) }) },
      align: { default: null, parseHTML: (el) => el.getAttribute("data-align"), renderHTML: (a) => (a.align ? { "data-align": a.align } : {}) },
      /** Soldan girinti (px; tblInd). Yalnız sola yaslı tabloda uygulanır. */
      indent: {
        default: 0,
        parseHTML: (el) => Math.round(Number(el.getAttribute("data-indent")) || 0),
        renderHTML: (a) => (a.indent ? { "data-indent": String(a.indent), ...(a.align && a.align !== "left" ? {} : { style: `margin-left:${a.indent}px` }) } : {}),
      },
      /** Başlık Satırlarını Yinele: ilk satır her sayfanın başında yinelenir. */
      repeatHeader: { default: false, parseHTML: (el) => el.getAttribute("data-repeat-header") === "true", renderHTML: (a) => (a.repeatHeader ? { "data-repeat-header": "true" } : {}) },
      /** Varsayılan hücre kenar boşlukları (pt; null: 0 / 0,19 cm). */
      cellMargins: {
        default: null,
        parseHTML: (el) => parseCellMargins(el.getAttribute("data-cpad")),
        renderHTML: (a) => (a.cellMargins ? { "data-cpad": cellMarginsAttr(a.cellMargins as CellMargins), style: `--bl-cpad:${cellPadCss(a.cellMargins as CellMargins)}` } : {}),
      },
    };
  },
}).configure({
  resizable: true,
  cellMinWidth: 30,
  lastColumnResizable: true,
  View: BelgeTableView as never,
  HTMLAttributes: { class: "bl-table" },
});

const cellAttrs = {
  background: {
    default: null,
    parseHTML: (el: HTMLElement) => el.getAttribute("data-bg") || el.style.backgroundColor || null,
    renderHTML: (a: Record<string, unknown>) => (a.background ? { "data-bg": a.background, style: `background-color:${a.background}` } : {}),
  },
  verticalAlign: {
    default: null,
    parseHTML: (el: HTMLElement) => el.style.verticalAlign || null,
    renderHTML: (a: Record<string, unknown>) => (a.verticalAlign ? { style: `vertical-align:${a.verticalAlign}` } : {}),
  },
};

export const BelgeTableCell = TableCell.extend({
  addAttributes() {
    return { ...this.parent?.(), ...cellAttrs };
  },
});

export const BelgeTableHeader = TableHeader.extend({
  addAttributes() {
    return { ...this.parent?.(), ...cellAttrs };
  },
});

export const BelgeTableRow = TableRow;
