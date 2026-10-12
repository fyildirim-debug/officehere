/** Slayt alanında metin düzenleyicinin (şekil ya da tablo hücresi) yerleştirilmesi. */
import { useCallback } from "react";
import type { Shape, SpShape, TableShape, TextBody } from "../model/types";
import type { TextRenderCtx } from "../render/textHtml";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";
import { TextEditor } from "./TextEditor";

export function useEditorRenderer(actions: Actions, getState: () => AppState) {
  return useCallback(
    (shape: SpShape | TableShape, trc: TextRenderCtx, cell?: [number, number]) => {
      const st = getState();
      const ed = st.editing;
      const body: TextBody | undefined = shape.kind === "sp" ? shape.text : cell ? shape.rows[cell[0]]?.cells[cell[1]]?.text : undefined;
      if (!body || !ed) return null;
      const key = `${shape.id}:${cell?.join(",") ?? ""}`;
      return (
        <TextEditor
          key={key}
          body={body}
          ctx={trc}
          initial={{ selectAll: ed.selectAll, replaceWith: ed.replaceWith, point: ed.point, range: ed.range }}
          spellcheck={st.spellcheck}
          lang={st.proofLang}
          maxLevel={shape.kind === "sp" && shape.ph ? 4 : 8}
          onChange={(nb, merge) => {
            const i = actions.s.slideIdx;
            actions.store.commit(
              (p) => {
                const sl = p.slides[i];
                const update = (s: Shape): Shape => {
                  if (s.id === shape.id) {
                    if (s.kind === "sp") return { ...s, text: nb };
                    if (s.kind === "table" && cell)
                      return { ...s, rows: s.rows.map((r, ri) => (ri !== cell[0] ? r : { ...r, cells: r.cells.map((c, ci) => (ci === cell[1] ? { ...c, text: nb } : c)) })) };
                  }
                  if (s.kind === "group") return { ...s, children: s.children.map(update) };
                  return s;
                };
                const slides = p.slides.slice();
                slides[i] = { ...sl, shapes: sl.shapes.map(update) };
                return { ...p, slides };
              },
              { merge: merge ? `${merge}:${key}` : undefined },
            );
            if (ed.replaceWith !== undefined || ed.selectAll || ed.point || ed.range) actions.store.set({ editing: { shapeId: ed.shapeId, cell: ed.cell } });
          }}
          onExit={() => actions.stopEditing()}
          register={(h) => {
            actions.textHandle = h;
          }}
          onSelectionChange={() => actions.store.set((s) => ({ textTick: s.textTick + 1 }))}
          onTab={
            cell && shape.kind === "table"
              ? (shift) => {
                  const nCols = shape.cols.length;
                  const nRows = shape.rows.length;
                  let idx = cell[0] * nCols + cell[1] + (shift ? -1 : 1);
                  if (idx < 0) idx = 0;
                  if (idx >= nRows * nCols) {
                    // Son hücrede Tab: yeni satır ekle.
                    actions.updateShapes((s) =>
                      s.kind === "table" ? { ...s, rows: [...s.rows, { h: s.rows[s.rows.length - 1].h, cells: s.cols.map(() => ({ text: { paragraphs: [{ runs: [], level: 0, props: {} }] } })) }], h: s.h + s.rows[s.rows.length - 1].h } : s,
                    [shape.id]);
                  }
                  actions.store.set({ editing: { shapeId: shape.id, cell: [Math.floor(idx / nCols), idx % nCols], selectAll: true } });
                  return true;
                }
              : undefined
          }
          onUndo={() => actions.store.undo()}
          onRedo={() => actions.store.redo()}
          onCopyRich={(paras) => {
            actions.store.clipboard = { kind: "text", paragraphs: paras };
          }}
          getRichClipboard={() => actions.store.clipboard?.paragraphs ?? null}
        />
      );
    },
    [actions],
  );
}
