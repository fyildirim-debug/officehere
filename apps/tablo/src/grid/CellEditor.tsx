/**
 * Hücre içi düzenleyici. Düzenleme yokken görünmez bir "tuş tuzağı"
 * olarak etkin hücrede durur: yazılan ilk karakter (IME dahil) düzenlemeyi
 * başlatır. Formül başvuruları renkli gösterilir (saydam metin + ayna katman).
 */
import { useEffect, useLayoutEffect, useMemo, useRef, useState, type RefObject } from "react";
import { fontStack } from "@officehere/core";
import { findRefSpans, TR_SYNTAX, type FnDef } from "@officehere/formula";
import { REF_COLORS, type Controller } from "../controller";
import { fontCss } from "../model/styles";
import { rect } from "../model/rect";
import { jumpCell, stepCell } from "../model/selection";
import type { GridLayout } from "./layout";
import { analyzeFormula, hintParts, suggestFunctions } from "./formulaAssist";

export function ColoredText({ text }: { text: string }) {
  if (!text.startsWith("=")) return <>{text + "​"}</>;
  const spans = findRefSpans(text, TR_SYNTAX);
  const out: React.ReactNode[] = [];
  let pos = 0;
  spans.forEach((s, i) => {
    if (s.start > pos) out.push(text.slice(pos, s.start));
    out.push(
      <span key={i} style={{ color: REF_COLORS[i % REF_COLORS.length] }}>
        {text.slice(s.start, s.end)}
      </span>,
    );
    pos = s.end;
  });
  out.push(text.slice(pos) + "​");
  return <>{out}</>;
}

/** Ok tuşlarıyla işaret kipinde başvuru hareketi. */
export function pointMove(ctl: Controller, dr: number, dc: number, extend: boolean, jump: boolean): void {
  const e = ctl.edit!;
  const sheet = ctl.wb.sheetById(e.pointSheetId ?? ctl.sheet.id) ?? ctl.sheet;
  const fromCorner = e.pointCorner ?? { row: e.row, col: e.col };
  const anchor = e.pointAnchor ?? fromCorner;
  const base = e.pointSpan ? fromCorner : { row: e.row, col: e.col };
  const next = jump ? jumpCell(ctl.wb, sheet, base, dr, dc) : stepCell(sheet, base, dr, dc);
  if (extend && e.pointSpan) ctl.pointTo(rect(anchor.row, anchor.col, next.row, next.col), sheet, anchor, next);
  else ctl.pointTo(rect(next.row, next.col), sheet, next, next);
  ctl.scrollIntoView(next.row, next.col);
}

interface AssistState {
  items: FnDef[];
  index: number;
  start: number;
  end: number;
}

/** Otomatik tamamlama ve bağımsız değişken ipucu katmanı. */
export function useFormulaAssist(ctl: Controller, active: boolean) {
  const e = ctl.edit;
  const [dismissed, setDismissed] = useState<string | null>(null);
  const info = useMemo(() => (e && active ? analyzeFormula(e.text, e.caret) : {}), [e?.text, e?.caret, active]);
  const items = useMemo(() => (info.prefix ? suggestFunctions(info.prefix.text) : []), [info.prefix?.text]);
  const [index, setIndex] = useState(0);
  useEffect(() => setIndex(0), [info.prefix?.text]);
  const key = `${info.prefix?.start}:${info.prefix?.text}`;
  const assist: AssistState | null = info.prefix && items.length && dismissed !== key ? { items, index: Math.min(index, items.length - 1), start: info.prefix.start, end: info.prefix.end } : null;
  const insert = (def: FnDef) => {
    if (!e || !assist) return;
    const text = e.text.slice(0, assist.start) + def.tr + "(" + e.text.slice(assist.end);
    ctl.updateEdit(text, assist.start + def.tr.length + 1);
  };
  return { assist, call: info.call, setIndex, insert, dismiss: () => setDismissed(key) };
}

export function AssistPopup({ assist, x, y, onPick, onHover }: { assist: AssistState; x: number; y: number; onPick: (d: FnDef) => void; onHover: (i: number) => void }) {
  const cur = assist.items[assist.index];
  return (
    <div className="tb-assist" style={{ left: x, top: y }} onMouseDown={(e) => e.preventDefault()}>
      <div className="tb-assist-list">
        {assist.items.map((d, i) => (
          <div key={d.name} className={`tb-assist-item${i === assist.index ? " is-active" : ""}`} onMouseEnter={() => onHover(i)} onDoubleClick={() => onPick(d)}>
            <span className="tb-assist-fx">fx</span>
            {d.tr}
          </div>
        ))}
      </div>
      {cur && <div className="tb-assist-desc">{cur.desc}</div>}
    </div>
  );
}

export function ArgHint({ def, argIndex, x, y }: { def: FnDef; argIndex: number; x: number; y: number }) {
  const p = hintParts(def, argIndex);
  return (
    <div className="tb-arghint" style={{ left: x, top: y }}>
      <span className="tb-arghint-name">{p.name}</span>(
      {p.args.map((a, i) => (
        <span key={i}>
          {i > 0 && "; "}
          <span className={a.active ? "is-active" : ""}>{a.text}</span>
        </span>
      ))}
      )
    </div>
  );
}

/** Düzenleme tuşlarını işler; true dönerse varsayılan engellenir. */
export function handleEditKey(ctl: Controller, ev: React.KeyboardEvent<HTMLTextAreaElement>, assistApi: ReturnType<typeof useFormulaAssist>): boolean {
  const e = ctl.edit;
  if (!e) return false;
  const ta = ev.currentTarget;
  const { assist } = assistApi;
  if (assist) {
    if (ev.key === "ArrowDown" || ev.key === "ArrowUp") {
      assistApi.setIndex((assist.index + (ev.key === "ArrowDown" ? 1 : assist.items.length - 1)) % assist.items.length);
      return true;
    }
    if (ev.key === "Tab") {
      assistApi.insert(assist.items[assist.index]);
      return true;
    }
    if (ev.key === "Escape") {
      assistApi.dismiss();
      return true;
    }
  }
  switch (ev.key) {
    case "Enter":
      if (ev.altKey) {
        const s = ta.selectionStart;
        const text = e.text.slice(0, s) + "\n" + e.text.slice(ta.selectionEnd);
        ctl.updateEdit(text, s + 1);
        requestAnimationFrame(() => ta.setSelectionRange(s + 1, s + 1));
        return true;
      }
      if (ctl.commitEdit(ev.ctrlKey ? null : ev.shiftKey ? "up" : "down", ev.ctrlKey)) ctl.focusGrid();
      return true;
    case "Tab":
      if (ctl.commitEdit(ev.shiftKey ? "left" : "right")) ctl.focusGrid();
      return true;
    case "Escape":
      ctl.cancelEdit();
      ctl.focusGrid();
      return true;
    case "F2":
      ctl.setEditMode(e.mode === "enter" ? "edit" : "enter");
      return true;
    case "F4":
      ctl.toggleAbsolute();
      return true;
    case "ArrowUp":
    case "ArrowDown":
    case "ArrowLeft":
    case "ArrowRight": {
      if (e.mode === "edit" || e.inBar) return false;
      const dr = ev.key === "ArrowDown" ? 1 : ev.key === "ArrowUp" ? -1 : 0;
      const dc = ev.key === "ArrowRight" ? 1 : ev.key === "ArrowLeft" ? -1 : 0;
      if (ctl.canPoint()) {
        pointMove(ctl, dr, dc, ev.shiftKey, ev.ctrlKey);
        return true;
      }
      if (ctl.commitEdit(null)) {
        ctl.move(dr, dc, false, false);
        ctl.focusGrid();
      }
      return true;
    }
  }
  return false;
}

interface Props {
  ctl: Controller;
  layout: GridLayout | null;
  inputRef: RefObject<HTMLTextAreaElement | null>;
  onGridKey: (e: React.KeyboardEvent<HTMLTextAreaElement>) => void;
}

export function CellEditor({ ctl, layout, inputRef, onGridKey }: Props) {
  const e = ctl.edit;
  const sheet = ctl.sheet;
  const editingHere = !!e && e.sheetId === sheet.id && !!layout;
  const composing = useRef(false);
  const mirrorRef = useRef<HTMLDivElement>(null);
  const assistApi = useFormulaAssist(ctl, editingHere && !e?.inBar);
  const [size, setSize] = useState({ w: 0, h: 0 });

  // Değeri denetleyiciyle eşitle (imleç konumu korunarak)
  useLayoutEffect(() => {
    const ta = inputRef.current;
    if (!ta) return;
    const want = editingHere ? e!.text : "";
    if (ta.value !== want && !composing.current) {
      ta.value = want;
      if (editingHere && document.activeElement === ta) ta.setSelectionRange(e!.caret, e!.caret);
    } else if (editingHere && document.activeElement === ta && ta.selectionStart !== e!.caret && e!.pointSpan) {
      ta.setSelectionRange(e!.caret, e!.caret);
    }
  });

  const pos = e && editingHere ? { row: e.row, col: e.col } : ctl.sel.active;
  const merge = sheet.mergeAt(pos.row, pos.col);
  const cellRect = layout ? (merge ? layout.cellRect(merge.r1, merge.c1, merge.r2, merge.c2) : layout.cellRect(pos.row, pos.col)) : { x: 0, y: 0, w: 0, h: 0 };
  const style = ctl.wb.styleAt(sheet, pos.row, pos.col);
  const z = layout?.z ?? 1;
  const font = fontCss(style, z, fontStack);
  const lineH = (style.size ?? 11) * (96 / 72) * z * 1.2;

  // Metin genişliğine göre büyüme
  useLayoutEffect(() => {
    const m = mirrorRef.current;
    if (!m || !editingHere || !layout) return;
    const maxW = Math.max(cellRect.w, layout.width - cellRect.x - 4);
    m.style.width = "auto";
    m.style.height = "auto";
    m.style.right = "auto";
    m.style.whiteSpace = style.wrap ? "pre-wrap" : "pre";
    let w = Math.max(cellRect.w, m.scrollWidth + 10);
    if (style.wrap) w = cellRect.w;
    if (w > maxW) w = maxW;
    m.style.width = `${w}px`;
    m.style.whiteSpace = "pre-wrap";
    const h = Math.max(cellRect.h, m.scrollHeight + 2);
    m.style.width = "";
    m.style.height = "";
    m.style.right = "";
    m.style.whiteSpace = "";
    if (Math.abs(w - size.w) > 0.5 || Math.abs(h - size.h) > 0.5) setSize({ w, h });
  });

  const visible = editingHere && !!layout;
  const boxStyle: React.CSSProperties = visible
    ? {
        left: cellRect.x - 1,
        top: cellRect.y - 1,
        width: (size.w || cellRect.w) + 2,
        height: Math.min((size.h || cellRect.h) + 2, layout!.height - cellRect.y),
        font,
        lineHeight: `${lineH}px`,
        background: style.fill ?? "#FFFFFF",
        color: style.color ?? "#000000",
        fontWeight: style.bold ? 700 : 400,
        fontStyle: style.italic ? "italic" : "normal",
        textDecoration: style.underline ? "underline" : undefined,
      }
    : { left: cellRect.x, top: cellRect.y, width: 4, height: 4 };

  return (
    <>
      <div className={`tb-editor${visible ? " is-visible" : ""}`} style={boxStyle}>
        <div ref={mirrorRef} className="tb-editor-mirror" style={visible ? { font, lineHeight: `${lineH}px` } : { display: "none" }} aria-hidden="true">
          {visible && <ColoredText text={e!.text} />}
        </div>
        <textarea
          ref={inputRef}
          className="tb-editor-input"
          spellCheck={false}
          autoComplete="off"
          aria-label="Hücre düzenleyici"
          style={visible ? { font, lineHeight: `${lineH}px`, caretColor: style.color ?? "#000" } : undefined}
          readOnly={!!e?.inBar}
          onCompositionStart={() => {
            composing.current = true;
            if (!ctl.edit) ctl.startEdit("");
          }}
          onCompositionEnd={(ev) => {
            composing.current = false;
            const ta = ev.currentTarget;
            ctl.updateEdit(ta.value, ta.selectionStart);
          }}
          onInput={(ev) => {
            const ta = ev.currentTarget;
            if (composing.current) return;
            if (!ctl.edit) {
              const text = ta.value;
              ctl.startEdit(text);
              if (ctl.edit) ctl.updateEdit(text, text.length);
              return;
            }
            ctl.updateEdit(ta.value, ta.selectionStart);
          }}
          onSelect={(ev) => {
            const ta = ev.currentTarget;
            if (ctl.edit && !ctl.edit.inBar && ta.selectionStart !== ctl.edit.caret) {
              ctl.edit = { ...ctl.edit, caret: ta.selectionStart, pointSpan: ta.selectionStart === ctl.edit.pointSpan?.end ? ctl.edit.pointSpan : undefined };
              ctl.emit();
            }
          }}
          onKeyDown={(ev) => {
            if (composing.current) return;
            if (ctl.edit && !ctl.edit.inBar) {
              if (handleEditKey(ctl, ev, assistApi)) ev.preventDefault();
              return;
            }
            onGridKey(ev);
          }}
        />
      </div>
      {visible && assistApi.assist && !e!.inBar && (
        <AssistPopup assist={assistApi.assist} x={cellRect.x} y={cellRect.y + (size.h || cellRect.h) + 2} onPick={(d) => assistApi.insert(d)} onHover={(i) => assistApi.setIndex(i)} />
      )}
      {visible && !assistApi.assist && assistApi.call && !e!.inBar && <ArgHint def={assistApi.call.def} argIndex={assistApi.call.argIndex} x={cellRect.x} y={cellRect.y + (size.h || cellRect.h) + 3} />}
    </>
  );
}
