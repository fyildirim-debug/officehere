// RTF tabloları: \trowd/\cellx satır tanımları, \cell/\row ile biten hücreler → ızgaralı tablo düğümü.
import type { JSONContent } from "@tiptap/core";
import type { FlowItem } from "../common";

export interface RowDef {
  /** Hücrelerin sağ kenarları (twip). */
  cellx: number[];
  /** Yatay birleştirme: 1 ilk, 2 devam. */
  merge: number[];
  /** Dikey birleştirme: 1 ilk, 2 devam. */
  vmerge: number[];
  bg: (string | null)[];
  /** Satırın sol kenarı (\trleft, twip) ve hizası (\trqc, \trqr). */
  left?: number;
  align?: "center" | "right" | null;
  /** Başlık satırı olarak her sayfada yinelenir (\trhdr). */
  header?: boolean;
}

interface Row {
  cells: JSONContent[][];
  def: RowDef;
}

const MAX_COLS = 64;

export class RtfTables {
  private rows: Row[] = [];
  private cells: JSONContent[][] = [];
  private cur: JSONContent[] = [];
  private lastDef: RowDef = { cellx: [], merge: [], vmerge: [], bg: [] };
  private def: RowDef = { cellx: [], merge: [], vmerge: [], bg: [] };
  private cellDef = { merge: 0, vmerge: 0, bg: 0 };

  /** Satır/hücre tanımı sözcükleri (\trowd, \cellx, birleştirme, gölge) ve \row; işlendiyse true. */
  word(w: string, param: number | null, color: (i: number) => string | null): boolean {
    const d = this.cellDef;
    switch (w) {
      case "trowd":
        this.def = { cellx: [], merge: [], vmerge: [], bg: [] };
        this.cellDef = { merge: 0, vmerge: 0, bg: 0 };
        return true;
      case "trleft":
        this.def.left = param ?? 0;
        return true;
      case "trhdr":
        this.def.header = true;
        return true;
      case "trql":
      case "trqc":
      case "trqr":
        this.def.align = w === "trqc" ? "center" : w === "trqr" ? "right" : null;
        return true;
      case "clmgf":
      case "clmrg":
        d.merge = w === "clmgf" ? 1 : 2;
        return true;
      case "clvmgf":
      case "clvmrg":
        d.vmerge = w === "clvmgf" ? 1 : 2;
        return true;
      case "clcbpat":
        d.bg = param ?? 0;
        return true;
      case "cellx":
        this.def.cellx.push(param ?? 0);
        this.def.merge.push(d.merge);
        this.def.vmerge.push(d.vmerge);
        this.def.bg.push(d.bg ? color(d.bg) : null);
        this.cellDef = { merge: 0, vmerge: 0, bg: 0 };
        return true;
      case "row":
      case "nestrow":
        this.endRow(this.def);
        return true;
    }
    return false;
  }

  addBlock(n: JSONContent) {
    this.cur.push(n);
  }

  /** Açık tablo yok mu (satır, hücre, blok). */
  isEmpty(): boolean {
    return !this.rows.length && !this.cells.length && !this.cur.length;
  }

  hasBlocks(): boolean {
    return this.cur.length > 0;
  }

  endCell() {
    this.cells.push(this.cur.length ? this.cur : [{ type: "paragraph" }]);
    this.cur = [];
  }

  endRow(def: RowDef) {
    if (this.cur.length) this.endCell();
    if (this.cells.length) this.rows.push({ cells: this.cells, def: { ...def, cellx: [...def.cellx], merge: [...def.merge], vmerge: [...def.vmerge], bg: [...def.bg] } });
    this.lastDef = def;
    this.cells = [];
  }

  /** Açık tabloyu akışa ekler (tablo dışı paragraf geldiğinde ya da belge sonunda). */
  close(items: FlowItem[]) {
    if (this.cur.length || this.cells.length) this.endRow(this.lastDef);
    if (!this.rows.length) return;
    items.push({ node: this.build() });
    this.rows = [];
  }

  private build(): JSONContent {
    // Her hücrenin sol/sağ kenarı; tanımsız hücrelere eşit genişlik.
    type Cell = { left: number; right: number; blocks: JSONContent[]; merge: number; vmerge: number; bg: string | null; rowspan: number; dropped: boolean };
    // Soldan girinti: ilk satırın \trleft değeri (eksi değer, Word'ün hücre boşluğu payıdır; yok sayılır).
    const indentTw = Math.max(0, this.rows[0]?.def.left ?? 0);
    const grid: Cell[][] = this.rows.map((r) => {
      const out: Cell[] = [];
      let left = indentTw;
      r.cells.slice(0, MAX_COLS).forEach((blocks, i) => {
        const right = r.def.cellx[i] !== undefined && r.def.cellx[i] > left ? r.def.cellx[i] : left + 1440;
        out.push({ left, right, blocks, merge: r.def.merge[i] ?? 0, vmerge: r.def.vmerge[i] ?? 0, bg: r.def.bg[i] ?? null, rowspan: 1, dropped: false });
        left = right;
      });
      return out;
    });
    // Yatay birleştirme: devam hücresi öncekine katılır.
    for (const row of grid) {
      for (let i = row.length - 1; i > 0; i--) {
        if (row[i].merge === 2) {
          row[i - 1].right = row[i].right;
          row[i - 1].blocks.push(...row[i].blocks.filter((b) => b.content?.length));
          row.splice(i, 1);
        }
      }
    }
    // Dikey birleştirme: aynı sol kenarlı üst hücrenin satır aralığı büyür.
    grid.forEach((row, r) => {
      for (const c of row) {
        if (c.vmerge !== 2) continue;
        for (let k = r - 1; k >= 0; k--) {
          const up = grid[k].find((x) => Math.abs(x.left - c.left) < 20 && !x.dropped);
          if (up) {
            up.rowspan++;
            const extra = c.blocks.filter((b) => b.content?.length);
            if (extra.length) up.blocks.push(...extra);
            c.dropped = true;
            break;
          }
          const covered = grid[k].find((x) => Math.abs(x.left - c.left) < 20 && x.dropped);
          if (!covered) break;
        }
      }
    });
    // Ortak sütun ızgarası: tüm kenarların birleşimi.
    const edges = [...new Set(grid.flatMap((row) => row.flatMap((c) => [c.left, c.right])))].sort((a, b) => a - b).slice(0, MAX_COLS + 1);
    const colW = edges.slice(1).map((e, i) => Math.max(10, Math.round((e - edges[i]) / 15)));
    const idx = (x: number) => {
      let best = 0;
      edges.forEach((e, i) => {
        if (Math.abs(e - x) < Math.abs(edges[best] - x)) best = i;
      });
      return best;
    };
    const total = colW.length;
    const rows: JSONContent[] = grid.map((row) => {
      const cells: JSONContent[] = [];
      let used = 0;
      for (const c of row) {
        const a = idx(c.left);
        const b = Math.max(a + 1, idx(c.right));
        const span = Math.max(1, Math.min(b, total) - a);
        used = Math.max(used, a + span);
        if (c.dropped) continue;
        cells.push({
          type: "tableCell",
          attrs: { colspan: span, rowspan: c.rowspan, colwidth: colW.slice(a, a + span), background: c.bg },
          content: c.blocks.length ? c.blocks : [{ type: "paragraph" }],
        });
      }
      // Kısa satırlar boş hücreyle tamamlanır.
      for (let k = used; k < total; k++) cells.push({ type: "tableCell", attrs: { colspan: 1, rowspan: 1, colwidth: [colW[k]] }, content: [{ type: "paragraph" }] });
      return { type: "tableRow", content: cells };
    });
    const align = this.rows[0]?.def.align ?? null;
    const indent = align ? 0 : Math.round(indentTw / 15);
    const repeatHeader = !!this.rows[0]?.def.header;
    return { type: "table", attrs: { tableStyle: "grid", headerRow: true, ...(align ? { align } : {}), ...(indent ? { indent } : {}), ...(repeatHeader ? { repeatHeader } : {}) }, content: rows.filter((r) => r.content?.length) };
  }
}
