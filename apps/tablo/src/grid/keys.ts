/** Düzenleme yokken ızgara klavye kısayolları (Excel ile aynı). */
import type { Controller } from "../controller";
import * as F from "../actions/format";
import * as E from "../actions/edit";
import * as D from "../actions/data";
import { MAX_COLS, MAX_ROWS } from "../model/rect";
import { validationList } from "../controller";
import { pivotAt, refreshAll, refreshPivot } from "../pivot/pivotOps";
import { chartOnNewSheet, quickChart } from "../actions/objects";
import { handleObjectKey } from "../actions/objectClip";
import { groupSelection, ungroupSelection } from "../actions/outline";
import { flashFill } from "../actions/whatIf";

export interface GridKeyContext {
  pageRows: () => number;
  pageCols: () => number;
}

export function handleGridKey(ctl: Controller, ev: KeyboardEvent | React.KeyboardEvent, kc: GridKeyContext): boolean {
  const ctrl = ev.ctrlKey || ev.metaKey;
  const shift = ev.shiftKey;
  const alt = ev.altKey;
  const key = ev.key;
  const code = (ev as KeyboardEvent).code ?? "";
  const sheet = ctl.sheet;
  // Seçili grafik/resim/şekil: Sil, Esc, ok tuşları, yazma
  if (ctl.selObj && handleObjectKey(ctl, ev)) return true;

  switch (key) {
    case "ArrowUp":
    case "ArrowDown":
    case "ArrowLeft":
    case "ArrowRight": {
      // Alt+Shift+Sağ/Sol: Gruplandır / Grubu Çöz
      if (alt && shift && (key === "ArrowRight" || key === "ArrowLeft")) {
        if (key === "ArrowRight") groupSelection(ctl);
        else ungroupSelection(ctl);
        return true;
      }
      if (alt && key === "ArrowDown") {
        const { row, col } = ctl.sel.active;
        const v = sheet.validations.find((x) => x.rule.type === "list" && x.ranges.some((r) => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2));
        if (v && v.rule.showDropdown !== false) {
          ctl.listDropdown = true;
          ctl.emit();
        } else if (v) ctl.openDialog("validationList", { items: validationList(ctl.wb, sheet, v.rule.formula1 ?? "", row, col) });
        else ctl.openDialog("pickList");
        return true;
      }
      const dr = key === "ArrowDown" ? 1 : key === "ArrowUp" ? -1 : 0;
      const dc = key === "ArrowRight" ? 1 : key === "ArrowLeft" ? -1 : 0;
      ctl.move(dr, dc, shift, ctrl);
      return true;
    }
    case "Enter":
      if (alt) return false;
      // Kopyalama çerçevesi varken Enter seçili hücreye yapıştırır (seçim ilerlemez)
      if (ctl.clip) {
        E.pasteData(ctl, "", undefined);
        ctl.clip = null;
        ctl.status = "";
        ctl.emit();
        return true;
      }
      ctl.advance(shift ? "up" : "down");
      return true;
    case "Tab":
      ctl.advance(shift ? "left" : "right");
      return true;
    case "F2":
      if (shift) {
        ctl.openDialog("note");
        return true;
      }
      ctl.startEdit(undefined);
      return true;
    case "Delete":
      E.clearContents(ctl);
      return true;
    case "Backspace":
      // Excel: içerik boş metinle düzenlemeye girer; Esc'de eski içerik korunur
      ctl.startEdit("");
      return true;
    case "Escape":
      ctl.cancelModes();
      return true;
    case "PageDown":
    case "PageUp": {
      const d = key === "PageDown" ? 1 : -1;
      if (alt) ctl.move(0, d * kc.pageCols(), shift, false);
      else if (ctrl) {
        const i = ctl.wb.activeIndex + d;
        if (i >= 0 && i < ctl.wb.sheets.length) ctl.activateSheet(i);
      } else ctl.move(d * kc.pageRows(), 0, shift, false);
      return true;
    }
    case "Home":
      if (ctrl) {
        const fr = sheet.frozen.rows;
        const fc = sheet.frozen.cols;
        if (shift) {
          ctl.corner = { row: fr, col: fc };
          ctl.setSelection({ ...ctl.sel, ranges: [...ctl.sel.ranges.slice(0, -1), { r1: Math.min(fr, ctl.sel.anchor.row), c1: Math.min(fc, ctl.sel.anchor.col), r2: Math.max(fr, ctl.sel.anchor.row), c2: Math.max(fc, ctl.sel.anchor.col) }] });
        } else ctl.selectCell(fr, fc);
      } else ctl.selectCell(ctl.sel.active.row, sheet.frozen.cols);
      return true;
    case "End":
      if (ctrl) {
        const u = ctl.wb.usedRange(sheet);
        if (u) ctl.selectCell(u.r2, u.c2);
        return true;
      }
      return false;
    case "F4":
      if (ctrl) return false;
      // F4: son işlemi yinele (biçim uygula, satır ekle…)
      ctl.repeatLast();
      return true;
    case "F5": {
      // Alt+F5: PivotTable'ı yenile, Ctrl+Alt+F5: tümünü yenile
      if (alt) {
        if (ctrl) refreshAll(ctl);
        else {
          const pv = pivotAt(sheet, ctl.sel.active.row, ctl.sel.active.col);
          if (pv) refreshPivot(ctl, sheet, pv.id);
        }
        return true;
      }
      ctl.openDialog("goto");
      return true;
    }
    case "F9":
      ctl.wb.engine.recalc(!shift);
      ctl.emit();
      return true;
    case "F3":
      if (ctrl) {
        ctl.openDialog(shift ? "createNames" : "nameManager");
        return true;
      }
      if (shift) {
        ctl.openDialog("insertFunction");
        return true;
      }
      return false;
    case "F11":
      if (shift) {
        ctl.ops.addSheet(ctl.wb.activeIndex);
        return true;
      }
      // F11: verinin grafiğini ayrı grafik sayfasında oluştur
      if (!ctrl && !alt) {
        chartOnNewSheet(ctl);
        return true;
      }
      return false;
    case "F7":
      ctl.openDialog("spelling");
      return true;
    case "F10":
      if (shift) return false;
      return false;
    case "F1":
      // Alt+F1: seçili veriden hızlı grafik
      if (alt && !ctrl) {
        quickChart(ctl);
        return true;
      }
      return false;
    case "ContextMenu":
      return false;
  }

  if (ctrl) {
    // Fiziksel tuşa göre (klavye düzeninden bağımsız): Ctrl+Shift+~ Genel, Ctrl+Shift+@ (2) saat biçimi
    if (shift && !alt && code === "Backquote") {
      F.setNumFmt(ctl, "General");
      return true;
    }
    if (shift && !alt && code === "Digit2") {
      F.setNumFmt(ctl, "h:mm AM/PM");
      return true;
    }
    const k = code.startsWith("Key") ? code.slice(3) : code.startsWith("Digit") ? code.slice(5) : key.toUpperCase();
    switch (k) {
      case "A":
        if (shift) return false;
        E.selectAllSmart(ctl);
        return true;
      case "B":
      case "2":
        F.toggle(ctl, "bold");
        return true;
      case "I":
      case "3":
        if (shift && k === "3") {
          // Ctrl+Shift+3: gün.ay adı.yıl tarih biçimi (gg.aaa.yy)
          F.setNumFmt(ctl, "dd.mmm.yy");
          return true;
        }
        F.toggle(ctl, "italic");
        return true;
      case "U":
      case "4":
        if (shift && k === "4") {
          F.setNumFmt(ctl, '#,##0.00 "₺"');
          return true;
        }
        F.toggleUnderline(ctl);
        return true;
      case "5":
        if (shift) {
          F.percent(ctl);
          return true;
        }
        F.toggle(ctl, "strike");
        return true;
      case "1":
        if (shift) {
          F.setNumFmt(ctl, "#,##0.00");
          return true;
        }
        ctl.openDialog("formatCells");
        return true;
      case "6":
        if (shift) {
          F.setNumFmt(ctl, "0.00E+00");
          return true;
        }
        return false;
      case "9":
        hideSelection(ctl, "row", !shift);
        return true;
      case "0":
        hideSelection(ctl, "col", !shift);
        return true;
      case "D":
        E.fillDirection(ctl, "down");
        return true;
      case "E":
        // Ctrl+E: Akıllı Doldurma
        flashFill(ctl);
        return true;
      case "R":
        E.fillDirection(ctl, "right");
        return true;
      case "Z":
        E.undo(ctl);
        return true;
      case "Y":
        E.redo(ctl);
        return true;
      case "F":
        // Ctrl+Shift+F: Hücreleri Biçimlendir > Yazı Tipi
        if (shift) ctl.openDialog("formatCells", { tab: "font" });
        else ctl.openDialog("find");
        return true;
      case "H":
        ctl.openDialog("replace");
        return true;
      case "G":
        ctl.openDialog("goto");
        return true;
      case "K":
        ctl.openDialog("hyperlink");
        return true;
      case "L":
        if (shift) {
          D.toggleAutoFilter(ctl);
          return true;
        }
        ctl.openDialog("createTable");
        return true;
      case "T":
        ctl.openDialog("createTable");
        return true;
      case "Backquote":
      case "`":
      case '"':
        // Ctrl+Shift+~: Genel sayı biçimi; Ctrl+`: formülleri göster
        if (shift && code === "Backquote") {
          F.setNumFmt(ctl, "General");
          return true;
        }
        sheet.showFormulas = !sheet.showFormulas;
        ctl.emit();
        return true;
      case "Quote":
      case "'":
        // Ctrl+' üstteki formülü, Ctrl+Shift+" üstteki değeri kopyalar
        E.copyFromAbove(ctl, shift);
        return true;
      case "Semicolon":
      case ";":
      case ":":
      case ",":
      case "Comma":
        // Fiziksel tuş (e.code) esas alınır: ABD düzeninde Ctrl+Shift+; ":" üretir
        E.insertNow(ctl, key === ":" || (shift && (code === "Semicolon" || k === ";")) ? "time" : "date");
        return true;
      case "Minus":
      case "-":
        E.deleteCellsSmart(ctl);
        return true;
      case "Equal":
      case "+":
      case "=":
        if (shift || key === "+") {
          E.insertCellsSmart(ctl);
          return true;
        }
        return false;
      case "Space":
      case " ":
        if (shift) ctl.selectAll();
        else ctl.selectColumns(ctl.range.c1, ctl.range.c2);
        return true;
    }
    if (key === " ") {
      ctl.selectColumns(ctl.range.c1, ctl.range.c2);
      return true;
    }
    return false;
  }
  if (shift && key === " ") {
    ctl.selectRows(ctl.range.r1, ctl.range.r2);
    return true;
  }
  if (alt && (key === "=" || code === "Equal")) {
    E.autoSum(ctl);
    return true;
  }
  return false;
}

function hideSelection(ctl: Controller, axis: "row" | "col", hidden: boolean): void {
  const items: number[] = [];
  for (const r of ctl.sel.ranges) {
    const [a, b] = axis === "row" ? [r.r1, Math.min(r.r2, r.r1 + 100000)] : [r.c1, r.c2];
    for (let i = a; i <= b; i++) items.push(i);
  }
  if (!hidden) {
    // Göster: seçimin içindeki gizli öğeler
    const ax = axis === "row" ? ctl.sheet.rows : ctl.sheet.cols;
    const hiddenItems = ax.hiddenItems().filter((i) => ctl.sel.ranges.some((r) => (axis === "row" ? i >= r.r1 && i <= r.r2 : i >= r.c1 && i <= r.c2)));
    ctl.ops.setHidden(ctl.sheet, axis, hiddenItems, false);
    return;
  }
  ctl.ops.setHidden(ctl.sheet, axis, items, true);
  void MAX_ROWS;
  void MAX_COLS;
}
