/** Klavye kısayolları (odak alanına göre: slaytlar bölmesi, slayt alanı, metin düzenleme). */
import { useEffect, useRef } from "react";
import type { Actions } from "../state/actions";
import { applyRun, growFont, setParaProps } from "../state/textActions";
import { captureFormat, applyFormatPainter } from "../state/painter";

export interface KeyHandlers {
  save: () => void;
  saveAs: () => void;
  open: () => void;
  newDoc: () => void;
  print: () => void;
  showFromStart: () => void;
  showFromCurrent: () => void;
  find: () => void;
  replace: () => void;
  isBlocked: () => boolean;
}

function isFormField(t: EventTarget | null): boolean {
  const el = t as HTMLElement | null;
  if (!el?.tagName) return false;
  if (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.tagName === "SELECT") return true;
  return el.isContentEditable && !el.classList.contains("tb-edit");
}

export function useKeyboard(actions: Actions, h: KeyHandlers) {
  const ref = useRef(h);
  ref.current = h;
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.defaultPrevented || e.isComposing) return;
      const hd = ref.current;
      if (hd.isBlocked()) return;
      const s = actions.s;
      const mod = e.ctrlKey || e.metaKey;
      const key = e.key;
      const code = e.code;
      const k = code.startsWith("Key") ? code.slice(3) : code.startsWith("Digit") ? code.slice(5) : key;
      const inField = isFormField(e.target);
      const editingText = !!s.editing && (e.target as HTMLElement)?.classList?.contains("tb-edit");
      const run = (fn: () => void) => {
        e.preventDefault();
        e.stopPropagation();
        fn();
      };

      // Her yerde geçerli kısayollar.
      if (key === "F5") return run(() => (e.shiftKey ? hd.showFromCurrent() : hd.showFromStart()));
      if (key === "F12") return run(() => hd.saveAs());
      if (mod && !e.altKey) {
        if (k === "S") return run(() => (e.shiftKey ? hd.saveAs() : hd.save()));
        if (k === "O") return run(() => hd.open());
        if (k === "N") return run(() => hd.newDoc());
        if (k === "P") return run(() => hd.print());
        if (inField) return;
        if (k === "Z") return run(() => (e.shiftKey ? actions.store.redo() : actions.store.undo()));
        if (k === "Y") return run(() => actions.store.redo());
        if (k === "M") return run(() => actions.newSlide());
        if (k === "F") return run(() => hd.find());
        if (k === "H") return run(() => hd.replace());
        if (k === "B") return run(() => applyRun(actions, (p, _pp, eff) => ({ ...p, bold: !eff.bold })));
        if (k === "I") return run(() => applyRun(actions, (p, _pp, eff) => ({ ...p, italic: !eff.italic })));
        if (k === "U") return run(() => applyRun(actions, (p, _pp, eff) => ({ ...p, underline: !eff.underline })));
        if (k === "E" && !e.shiftKey) return run(() => setParaProps(actions, { align: "ctr" }));
        if (k === "L" && !e.shiftKey) return run(() => setParaProps(actions, { align: "l" }));
        if (k === "R" && !e.shiftKey) return run(() => setParaProps(actions, { align: "r" }));
        if (k === "J" && !e.shiftKey) return run(() => setParaProps(actions, { align: "just" }));
        if (key === "]" || (e.shiftKey && (key === ">" || code === "Period"))) return run(() => growFont(actions, 1));
        if (key === "[" || (e.shiftKey && (key === "<" || code === "Comma"))) return run(() => growFont(actions, -1));
        if (code === "IntlBackslash" && e.shiftKey) return run(() => growFont(actions, 1));
        if (code === "IntlBackslash") return run(() => growFont(actions, -1));
        if (key === " " || code === "Space")
          return run(() => applyRun(actions, (p) => ({ ...(p.field ? { field: p.field } : {}) })));
        if (e.shiftKey && k === "C") return run(() => captureFormat(actions, false));
        if (e.shiftKey && k === "V")
          return run(() => {
            const sel = actions.s.sel;
            for (const id of sel) applyFormatPainter(actions, id);
          });
      }
      if (inField || editingText) return;

      const focus = s.focus;
      if (focus === "slides") {
        if (key === "Delete" || key === "Backspace") return run(() => actions.deleteSlides());
        if (mod && k === "D") return run(() => actions.duplicateSlides());
        if (mod && k === "A") return run(() => actions.store.set({ slideSel: s.pres.slides.map((x) => x.id) }));
        if (key === "ArrowUp" || key === "ArrowLeft" || key === "PageUp") return run(() => actions.goToSlide(s.slideIdx - 1, e.shiftKey ? "extend" : "replace"));
        if (key === "ArrowDown" || key === "ArrowRight" || key === "PageDown") return run(() => actions.goToSlide(s.slideIdx + 1, e.shiftKey ? "extend" : "replace"));
        if (key === "Home") return run(() => actions.goToSlide(0));
        if (key === "End") return run(() => actions.goToSlide(s.pres.slides.length - 1));
        if (key === "Enter") return run(() => actions.newSlide());
        return;
      }
      if (focus === "notes") return;

      // Slayt alanı.
      const hasSel = s.sel.length > 0;
      if (mod && k === "A") return run(() => actions.selectAll());
      if (mod && k === "D") return run(() => (hasSel ? actions.duplicateSelection() : actions.duplicateSlides()));
      if (mod && k === "G") return run(() => (e.shiftKey ? actions.ungroupSelection() : actions.groupSelection()));
      if (key === "Escape") {
        if (s.insert) return run(() => actions.store.set({ insert: null }));
        if (s.cropId) return run(() => actions.store.set({ cropId: null }));
        if (s.ink) return run(() => actions.store.set({ ink: null }));
        if (s.painter) return run(() => actions.store.set({ painter: null }));
        if (hasSel) return run(() => actions.select([]));
        return;
      }
      if (key === "Tab") return run(() => actions.cycleSelection(e.shiftKey));
      if (!hasSel) {
        if (key === "ArrowUp" || key === "ArrowLeft" || key === "PageUp") return run(() => actions.goToSlide(s.slideIdx - 1));
        if (key === "ArrowDown" || key === "ArrowRight" || key === "PageDown") return run(() => actions.goToSlide(s.slideIdx + 1));
        if (key === "Home") return run(() => actions.goToSlide(0));
        if (key === "End") return run(() => actions.goToSlide(s.pres.slides.length - 1));
        return;
      }
      if (key === "Delete" || key === "Backspace") return run(() => actions.deleteSelection());
      if (key.startsWith("Arrow")) {
        const step = mod ? 1 : (96 / 2.54) * 0.2;
        const dx = key === "ArrowLeft" ? -step : key === "ArrowRight" ? step : 0;
        const dy = key === "ArrowUp" ? -step : key === "ArrowDown" ? step : 0;
        if (e.altKey) return run(() => actions.rotateSelection(key === "ArrowLeft" ? -15 : key === "ArrowRight" ? 15 : 0));
        return run(() => actions.moveSelection(dx, dy, "nudge"));
      }
      const shapes = actions.selectedShapes();
      if ((key === "Enter" || key === "F2") && shapes.length === 1) {
        const sh = shapes[0];
        if (sh.kind === "sp") return run(() => actions.startEditing({ shapeId: sh.id, selectAll: true }));
        if (sh.kind === "table") return run(() => actions.startEditing({ shapeId: sh.id, cell: [0, 0], selectAll: true }));
      }
      if (!mod && !e.altKey && key.length === 1 && shapes.length === 1 && shapes[0].kind === "sp") {
        const sh = shapes[0];
        if (sh.kind === "sp" && sh.geom !== "line" && sh.geom !== "straightConnector1") return run(() => actions.startEditing({ shapeId: sh.id, replaceWith: key }));
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [actions]);
}
