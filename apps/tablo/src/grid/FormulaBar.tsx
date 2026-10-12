/** Formül çubuğu: Ad Kutusu, İptal/Gir/İşlev Ekle düğmeleri ve genişletilebilir giriş. */
import { useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from "react";
import { Menu, MenuItem, Popover, tipProps } from "@officehere/ui";
import { addr } from "@officehere/formula";
import type { Controller } from "../controller";
import { rectCols, rectRows } from "../model/rect";
import { ArgHint, AssistPopup, ColoredText, handleEditKey, useFormulaAssist } from "./CellEditor";

export function FormulaBar({ ctl }: { ctl: Controller }) {
  useSyncExternalStore(ctl.subscribe, ctl.getVersion);
  const [nameText, setNameText] = useState("");
  const [nameFocus, setNameFocus] = useState(false);
  const [namesOpen, setNamesOpen] = useState(false);
  const nameRef = useRef<HTMLDivElement>(null);
  const taRef = useRef<HTMLTextAreaElement>(null);
  const e = ctl.edit;
  const inBar = !!e?.inBar;
  const assistApi = useFormulaAssist(ctl, inBar);
  const sheet = ctl.sheet;
  const r = ctl.range;
  const multi = rectRows(r) > 1 || rectCols(r) > 1;
  const nameBox = nameFocus ? nameText : ctl.nameBoxText();
  const value = e ? e.text : ctl.wb.editText(sheet, ctl.sel.active.row, ctl.sel.active.col);
  const spill = !e ? ctl.wb.engine.getSpill(sheet.id, ctl.sel.active.row, ctl.sel.active.col) : undefined;
  const spillChild = spill && (spill.anchorRow !== ctl.sel.active.row || spill.anchorCol !== ctl.sel.active.col);
  const shown = spillChild ? ctl.wb.editText(sheet, spill.anchorRow, spill.anchorCol) : value;

  useLayoutEffect(() => {
    const ta = taRef.current;
    if (!ta) return;
    if (ta.value !== shown) ta.value = shown;
    if (inBar && document.activeElement === ta && e && ta.selectionStart !== e.caret && e.pointSpan) ta.setSelectionRange(e.caret, e.caret);
  });

  useEffect(() => {
    if (inBar && document.activeElement !== taRef.current) taRef.current?.focus();
  }, [inBar]);

  const names = [...ctl.wb.engine.getNames().map((n) => n.name), ...ctl.wb.sheets.flatMap((s) => s.tables.map((t) => t.name))];

  return (
    <div className={`tb-fbar${ctl.formulaBarExpanded ? " is-expanded" : ""}`}>
      <div className="tb-namebox" ref={nameRef}>
        <input
          className="tb-namebox-input"
          value={nameBox}
          aria-label="Ad Kutusu"
          {...tipProps("Ad Kutusu")}
          onFocus={(ev) => {
            setNameFocus(true);
            setNameText(ctl.nameBoxText());
            requestAnimationFrame(() => ev.target.select());
          }}
          onBlur={() => setNameFocus(false)}
          onChange={(ev) => setNameText(ev.target.value)}
          onKeyDown={(ev) => {
            if (ev.key === "Enter") {
              const t = nameText.trim();
              if (!ctl.goTo(t)) {
                // Geçerli bir ad ise seçim için tanımla (Excel davranışı)
                if (/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(t) && !/^[A-Za-z]{1,3}\d+$/.test(t)) {
                  ctl.wb.engine.defineName(t, `=${quote(sheet.name)}!${absRange(ctl)}`);
                  ctl.emit();
                } else ctl.alert("Geçersiz bir başvuru ya da ad girdiniz.");
              }
              setNameFocus(false);
              ctl.focusGrid();
            } else if (ev.key === "Escape") {
              setNameFocus(false);
              ctl.focusGrid();
            }
          }}
        />
        <button type="button" className="tb-namebox-arrow" aria-label="Adlar" onMouseDown={(ev) => ev.preventDefault()} onClick={() => setNamesOpen(!namesOpen)}>
          <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden="true">
            <path d="M1 2.5l3 3 3-3" fill="none" stroke="currentColor" strokeWidth="1.1" />
          </svg>
        </button>
        <Popover open={namesOpen} onClose={() => setNamesOpen(false)} anchor={nameRef}>
          <Menu width={160}>
            {names.length ? names.map((n) => <MenuItem key={n} label={n} onClick={() => (setNamesOpen(false), ctl.goTo(n), ctl.focusGrid())} />) : <MenuItem label="(Tanımlı ad yok)" disabled />}
          </Menu>
        </Popover>
      </div>
      <div className="tb-fbar-sep" />
      <div className="tb-fbar-btns">
        <button type="button" className="tb-fbar-btn" disabled={!e} aria-label="İptal" {...tipProps("İptal")} onMouseDown={(ev) => ev.preventDefault()} onClick={() => (ctl.cancelEdit(), ctl.focusGrid())}>
          <svg width="12" height="12" viewBox="0 0 12 12" aria-hidden="true">
            <path d="M2 2l8 8M10 2l-8 8" stroke="currentColor" strokeWidth="1.4" />
          </svg>
        </button>
        <button type="button" className="tb-fbar-btn" disabled={!e} aria-label="Gir" {...tipProps("Gir")} onMouseDown={(ev) => ev.preventDefault()} onClick={() => ctl.commitEdit(null) && ctl.focusGrid()}>
          <svg width="13" height="12" viewBox="0 0 13 12" aria-hidden="true">
            <path d="M1.5 6.5l3.5 3.5 6.5-8" fill="none" stroke="currentColor" strokeWidth="1.4" />
          </svg>
        </button>
        <button type="button" className="tb-fbar-btn tb-fbar-fx" aria-label="İşlev Ekle" {...tipProps("İşlev Ekle", "Shift+F3")} onMouseDown={(ev) => ev.preventDefault()} onClick={() => ctl.openDialog("insertFunction")}>
          <i>fx</i>
        </button>
      </div>
      <div className="tb-fbar-input-wrap">
        <div className={`tb-fbar-mirror${spillChild ? " is-ghost" : ""}`} aria-hidden="true">
          <ColoredText text={shown} />
        </div>
        <textarea
          ref={taRef}
          className="tb-fbar-input"
          spellCheck={false}
          aria-label="Formül Çubuğu"
          defaultValue={shown}
          onFocus={() => {
            if (!ctl.edit) {
              if (spillChild) return;
              ctl.startEdit(undefined, true);
            } else if (!ctl.edit.inBar) {
              ctl.edit = { ...ctl.edit, inBar: true, mode: "edit" };
              ctl.emit();
            }
          }}
          onInput={(ev) => {
            const ta = ev.currentTarget;
            if (!ctl.edit) ctl.startEdit(undefined, true);
            ctl.updateEdit(ta.value, ta.selectionStart);
          }}
          onSelect={(ev) => {
            const ta = ev.currentTarget;
            if (ctl.edit?.inBar && ta.selectionStart !== ctl.edit.caret) {
              ctl.edit = { ...ctl.edit, caret: ta.selectionStart, pointSpan: ta.selectionStart === ctl.edit.pointSpan?.end ? ctl.edit.pointSpan : undefined };
              ctl.emit();
            }
          }}
          onKeyDown={(ev) => {
            if (!ctl.edit) return;
            if (ev.key.startsWith("Arrow") && !assistApi.assist) return;
            if (handleEditKey(ctl, ev, assistApi)) ev.preventDefault();
          }}
        />
        {inBar && assistApi.assist && <AssistPopup assist={assistApi.assist} x={6} y={28} onPick={(d) => assistApi.insert(d)} onHover={(i) => assistApi.setIndex(i)} />}
        {inBar && !assistApi.assist && assistApi.call && <ArgHint def={assistApi.call.def} argIndex={assistApi.call.argIndex} x={6} y={28} />}
      </div>
      <button type="button" className="tb-fbar-expand" aria-label={ctl.formulaBarExpanded ? "Formül Çubuğunu Daralt" : "Formül Çubuğunu Genişlet"} {...tipProps(ctl.formulaBarExpanded ? "Formül Çubuğunu Daralt" : "Formül Çubuğunu Genişlet", "Ctrl+Shift+U")} onClick={() => ((ctl.formulaBarExpanded = !ctl.formulaBarExpanded), ctl.emit())}>
        <svg width="10" height="10" viewBox="0 0 10 10" aria-hidden="true">
          {ctl.formulaBarExpanded ? <path d="M2 6.5l3-3 3 3" fill="none" stroke="currentColor" strokeWidth="1.2" /> : <path d="M2 3.5l3 3 3-3" fill="none" stroke="currentColor" strokeWidth="1.2" />}
        </svg>
      </button>
      {multi && !nameFocus && <span className="tb-sr-only">{addr(r.r1, r.c1)}</span>}
    </div>
  );
}

function quote(name: string): string {
  return /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

function absRange(ctl: Controller): string {
  const r = ctl.range;
  const a = (row: number, col: number) => addr(row, col).replace(/^([A-Z]+)(\d+)$/, "$$$1$$$2");
  return r.r1 === r.r2 && r.c1 === r.c2 ? a(r.r1, r.c1) : `${a(r.r1, r.c1)}:${a(r.r2, r.c2)}`;
}
