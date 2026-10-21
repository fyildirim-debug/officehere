// @vitest-environment jsdom
// Tablo satır yüksekliği: seçimin kapsadığı satırlar, Yükseklik, Satırları Dağıt ve satır görünümü.
import { afterEach, describe, expect, it, vi } from "vitest";
import { Editor, type JSONContent } from "@tiptap/core";
import { NodeSelection, TextSelection } from "@tiptap/pm/state";
import { CellSelection } from "@tiptap/pm/tables";
import { normalizeDoc } from "../../model/document";
import { schemaExtensions } from "../extensions";
import { MIN_ROW_PX, currentRowHeight, distributeRows, rowHeightsTr, rowPositions, selectedRows, setRowHeight } from "../tableRowHeight";

let editor: Editor | null = null;
afterEach(() => {
  editor?.destroy();
  editor = null;
  vi.restoreAllMocks();
});

const cell = (s: string): JSONContent => ({ type: "tableCell", attrs: { colspan: 1, rowspan: 1 }, content: [{ type: "paragraph", content: [{ type: "text", text: s }] }] });
const row = (s: string, attrs?: Record<string, unknown>): JSONContent => ({ type: "tableRow", ...(attrs ? { attrs } : {}), content: [cell(`${s}1`), cell(`${s}2`)] });
function make(rows: JSONContent[]) {
  const content = normalizeDoc({ type: "doc", content: [{ type: "section", content: [{ type: "table", content: rows }, { type: "paragraph" }] }] });
  editor = new Editor({ extensions: schemaExtensions(), content });
  return editor;
}
function tablePos(e: Editor): number {
  let at = -1;
  e.state.doc.descendants((n, p) => {
    if (at < 0 && n.type.name === "table") at = p;
    return at < 0;
  });
  return at;
}
/** Metnin bulunduğu hücreye imleç. */
function cursorIn(e: Editor, text: string) {
  let at = -1;
  e.state.doc.descendants((n, p) => {
    if (at < 0 && n.isText && n.text === text) at = p;
    return at < 0;
  });
  e.view.dispatch(e.state.tr.setSelection(TextSelection.create(e.state.doc, at + 1)));
}
const cellPos = (e: Editor, r: number, c: number) => {
  const tp = tablePos(e);
  return rowPositions(e.state.doc.nodeAt(tp)!, tp)[r] + 1 + [0, e.state.doc.nodeAt(rowPositions(e.state.doc.nodeAt(tp)!, tp)[r])!.child(0).nodeSize][c];
};
const heights = (e: Editor) => {
  const t = e.state.doc.nodeAt(tablePos(e))!;
  const out: (number | null)[] = [];
  t.forEach((r) => out.push(r.attrs.height as number | null));
  return out;
};

describe("Tablo satır yüksekliği", () => {
  it("seçimin kapsadığı satırlar: imleç, hücre seçimi ve tüm tablo", () => {
    const e = make([row("A"), row("B"), row("C")]);
    cursorIn(e, "B2");
    expect(selectedRows(e.state)).toMatchObject({ rows: [1], all: false });
    e.view.dispatch(e.state.tr.setSelection(CellSelection.create(e.state.doc, cellPos(e, 0, 1), cellPos(e, 1, 0))));
    expect(selectedRows(e.state)).toMatchObject({ rows: [0, 1], all: false });
    e.view.dispatch(e.state.tr.setSelection(NodeSelection.create(e.state.doc, tablePos(e))));
    expect(selectedRows(e.state)).toMatchObject({ rows: [0, 1, 2], all: true });
  });

  it("Yükseklik seçili satırlara yazılır; boş değer kaldırır, kural da düşer; en küçük yükseklik korunur", () => {
    const e = make([row("A"), row("B", { height: 30, heightRule: "exact" }), row("C")]);
    e.view.dispatch(e.state.tr.setSelection(CellSelection.create(e.state.doc, cellPos(e, 0, 0), cellPos(e, 1, 1))));
    expect(setRowHeight(e.view, 50)).toBe(true);
    expect(heights(e)).toEqual([50, 50, null]);
    // Kural korunur (yalnız yükseklik değişti).
    expect(e.state.doc.nodeAt(tablePos(e))!.child(1).attrs.heightRule).toBe("exact");
    setRowHeight(e.view, null);
    expect(heights(e)).toEqual([null, null, null]);
    expect(e.state.doc.nodeAt(tablePos(e))!.child(1).attrs.heightRule).toBeNull();
    const tr = rowHeightsTr(e.state.tr, tablePos(e), [2, null, 12.345]);
    // Değişmeyen satıra adım yazılmaz.
    expect(tr.steps).toHaveLength(2);
    e.view.dispatch(tr);
    expect(heights(e)).toEqual([MIN_ROW_PX, null, 12.35]);
  });

  it("satır görünümü yüksekliği <tr> öğesine yansıtır", () => {
    const e = make([row("A", { height: 44 }), row("B")]);
    const tp = tablePos(e);
    const [r0, r1] = rowPositions(e.state.doc.nodeAt(tp)!, tp).map((p) => e.view.nodeDOM(p) as HTMLElement);
    expect(r0.tagName).toBe("TR");
    expect([r0.style.height, r1.style.height]).toEqual(["44px", ""]);
    cursorIn(e, "B1");
    setRowHeight(e.view, 70);
    expect((e.view.nodeDOM(rowPositions(e.state.doc.nodeAt(tp)!, tp)[1]) as HTMLElement).style.height).toBe("70px");
  });

  it("Satırları Dağıt: hücre seçimi yoksa tüm satırlar ortalama yüksekliği alır; Yükseklik kutusu ölçüleni gösterir", () => {
    const e = make([row("A"), row("B"), row("C", { height: 60 })]);
    const H: Record<string, number> = { A: 20, B: 40, C: 60 };
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockImplementation(function (this: HTMLElement) {
      const h = this.tagName === "TR" ? (H[(this.textContent ?? "")[0]] ?? 0) : 0;
      return { x: 0, y: 0, left: 0, top: 0, right: 100, bottom: h, width: 100, height: h, toJSON: () => ({}) } as DOMRect;
    });
    cursorIn(e, "A1");
    expect(currentRowHeight(e.view)).toBe(20);
    cursorIn(e, "C1");
    expect(currentRowHeight(e.view)).toBe(60);
    expect(distributeRows(e.view)).toBe(true);
    expect(heights(e)).toEqual([40, 40, 40]);
  });
});
