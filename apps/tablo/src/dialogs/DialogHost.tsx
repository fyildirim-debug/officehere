/** İletişim kutusu yöneticisi: kimliğe göre ilgili iletişim kutusunu çizer. */
import { useEffect } from "react";
import type { Controller } from "../controller";
import type { FileActions } from "../FileView";
import { GoToSpecialDialog, SpellingDialog, ValidationErrorDialog } from "./Round3Dialogs";
import { GoalSeekDialog, GoalSeekStatusDialog, GroupAxisDialog, SubtotalDialog } from "./DataToolsDialogs";
import * as D from "../actions/data";
import * as M from "./MiscDialogs";
import { FormatCellsDialog } from "./FormatCellsDialog";
import { FindDialog } from "./FindDialog";
import { InsertFunctionDialog } from "./FunctionDialog";
import { CustomFilterDialog, DataValidationDialog, RemoveDuplicatesDialog, SeriesDialog, SortDialog, TextImportDialog, TextToColumnsDialog } from "./DataDialogs";
import { CondManagerDialog, CondRuleDialog } from "./CondDialogs";
import { CreateNamesDialog, DefineNameDialog, NameManagerDialog } from "./NameDialogs";
import { ChartDataDialog, ChartTypeDialog, MoveChartDialog } from "./ChartDialogs";
import { PageSetupDialog } from "./PageSetupDialog";
import { ShortcutsDialog, WorkbookStatsDialog } from "./HelpDialogs";
import { DialogErrorBoundary } from "./common";
import { DataTableDialog, ScenarioEditDialog, ScenarioManagerDialog, ScenarioSummaryDialog, ScenarioValuesDialog } from "./WhatIfDialogs";
import { AdvancedFilterDialog, InsertSlicerDialog, ProtectWorkbookDialog } from "./Round4Dialogs";
import { ErrorCheckDialog, ErrorOptionsDialog, EvaluateFormulaDialog } from "./AuditDialogs";
import { traceError } from "../actions/audit";
import { SparklineAxisDialog, SparklineDialog, SparklineHiddenDialog, SparklineItemDialog } from "../sparkline/SparkDialogs";
import { ChangePivotSourceDialog, CreatePivotDialog, MovePivotDialog, PivotOptionsDialog, ValueFieldDialog } from "../pivot/PivotDialogs";

export interface DialogState {
  id: string;
  arg?: unknown;
  /** Yığındaki benzersiz anahtar (aynı anda birden çok iletişim kutusu: ör. Bul + uyarı). */
  key?: number;
}

/** Dialog açmadan doğrudan çalışan eylemler. */
const ACTIONS: Record<string, (ctl: Controller) => void> = {
  toggleFilter: (ctl) => D.toggleAutoFilter(ctl),
  traceError: (ctl) => traceError(ctl),
};

const handled = new WeakSet<DialogState>();

/** İletişim kutusu yığını: en üstteki etkindir; uyarılar alttakini kapatmadan üstte açılır. */
export function DialogHost({ ctl, stack, onClose, files, onBackstage }: { ctl: Controller; stack: DialogState[]; onClose: (key: number | undefined) => void; files: FileActions; onBackstage?: (page: string) => void }) {
  // Odak iletişim kutusunun dışına kaçtıysa (ör. ızgaraya döndüyse) Esc yine en üstteki kutuyu kapatır
  useEffect(() => {
    if (!stack.length) return;
    const top = stack[stack.length - 1];
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== "Escape" || e.defaultPrevented) return;
      const t = e.target as HTMLElement | null;
      if (t?.closest?.(".oh-dialog, .oh-popover, .oh-menu")) return;
      if (!document.querySelector(".oh-dialog")) return;
      e.preventDefault();
      e.stopPropagation();
      onClose(top.key);
    };
    window.addEventListener("keydown", onKey, true);
    return () => window.removeEventListener("keydown", onKey, true);
  }, [stack, onClose]);
  return (
    <>
      {stack.map((st) => (
        <DialogErrorBoundary
          key={st.key}
          resetKey={st}
          onError={(e) => {
            onClose(st.key);
            ctl.alert(`İletişim kutusu açılamadı: ${e.message}`);
          }}
        >
          <OneDialog ctl={ctl} state={st} onClose={() => onClose(st.key)} files={files} onBackstage={onBackstage} />
        </DialogErrorBoundary>
      ))}
    </>
  );
}

function OneDialog({ ctl, state, onClose, files, onBackstage }: { ctl: Controller; state: DialogState; onClose: () => void; files: FileActions; onBackstage?: (page: string) => void }) {
  const id = state.id;
  useEffect(() => {
    // Eylem bir kez çalışır (React'in geliştirme kipindeki çift efekt çağrısında da)
    if (ACTIONS[id]) {
      if (handled.has(state)) return;
      handled.add(state);
      ACTIONS[id](ctl);
      onClose();
    } else if (id === "printPreview" || id === "about") {
      onClose();
      if (onBackstage) onBackstage(id === "about" ? "about" : "print");
      else if (id === "printPreview") files.print();
    }
  }, [state]); // eslint-disable-line react-hooks/exhaustive-deps
  // Ortak iletişim kutusu ilk denetimi odaklar; işaretli alan varsa ona geç
  useEffect(() => {
    const raf = requestAnimationFrame(() => {
      const boxes = document.querySelectorAll(".oh-dialog");
      const el = boxes[boxes.length - 1]?.querySelector<HTMLInputElement>("[data-autofocus]");
      if (!el) return;
      el.focus();
      if (typeof el.select === "function") el.select();
    });
    return () => cancelAnimationFrame(raf);
  }, [state]);
  if (ACTIONS[state.id] || state.id === "printPreview" || state.id === "about") return null;
  const p = { ctl, onClose, arg: state.arg };
  switch (state.id) {
    case "alert":
      return <M.AlertDialog {...p} />;
    case "validationError":
      return <ValidationErrorDialog {...p} />;
    case "goSpecial":
      return <GoToSpecialDialog {...p} />;
    case "spelling":
      return <SpellingDialog {...p} />;
    case "groupAxis":
      return <GroupAxisDialog {...p} />;
    case "subtotal":
      return <SubtotalDialog {...p} />;
    case "goalSeek":
      return <GoalSeekDialog {...p} />;
    case "scenarioManager":
      return <ScenarioManagerDialog {...p} />;
    case "scenarioEdit":
      return <ScenarioEditDialog {...p} />;
    case "scenarioValues":
      return <ScenarioValuesDialog {...p} />;
    case "scenarioSummary":
      return <ScenarioSummaryDialog {...p} />;
    case "dataTable":
      return <DataTableDialog {...p} />;
    case "advancedFilter":
      return <AdvancedFilterDialog {...p} />;
    case "protectWorkbook":
      return <ProtectWorkbookDialog {...p} />;
    case "insertSlicer":
      return <InsertSlicerDialog {...p} />;
    case "goalSeekStatus":
      return <GoalSeekStatusDialog {...p} />;
    case "confirm":
      return <M.ConfirmDialog {...p} />;
    case "formatCells":
      return <FormatCellsDialog {...p} />;
    case "insertFunction":
      return <InsertFunctionDialog {...p} />;
    case "find":
      return <FindDialog {...p} />;
    case "replace":
      return <FindDialog {...p} arg={{ ...((state.arg as object) ?? {}), replace: true }} />;
    case "goto":
      return <M.GoToDialog {...p} />;
    case "pasteSpecial":
      return <M.PasteSpecialDialog {...p} />;
    case "sort":
      return <SortDialog {...p} />;
    case "condRule":
      return <CondRuleDialog {...p} />;
    case "condManager":
      return <CondManagerDialog {...p} />;
    case "dataValidation":
      return <DataValidationDialog {...p} />;
    case "nameManager":
      return <NameManagerDialog {...p} />;
    case "defineName":
      return <DefineNameDialog {...p} />;
    case "createNames":
      return <CreateNamesDialog {...p} />;
    case "rowHeight":
    case "colWidth":
    case "defaultWidth":
      return <M.SizeDialog {...p} arg={state.id} />;
    case "insertCells":
      return <M.InsertDeleteCellsDialog {...p} arg="insert" />;
    case "deleteCells":
      return <M.InsertDeleteCellsDialog {...p} arg="delete" />;
    case "moveSheet":
      return <M.MoveSheetDialog {...p} />;
    case "deleteSheet":
      return <M.DeleteSheetDialog {...p} />;
    case "unhideSheet":
      return <M.UnhideSheetDialog {...p} />;
    case "hyperlink":
      return <M.HyperlinkDialog {...p} />;
    case "note":
      return <M.NoteDialog {...p} />;
    case "symbol":
      return <M.SymbolDialog {...p} />;
    case "zoom":
      return <M.ZoomDialog {...p} />;
    case "createTable":
      return <M.CreateTableDialog {...p} />;
    case "chartType":
      return <ChartTypeDialog {...p} />;
    case "chartData":
      return <ChartDataDialog {...p} />;
    case "moveChart":
      return <MoveChartDialog {...p} />;
    case "protect":
      return <M.ProtectDialog {...p} />;
    case "customFilter":
      return <CustomFilterDialog {...p} />;
    case "validationList":
    case "pickList":
      return <M.PickListDialog {...p} />;
    case "series":
      return <SeriesDialog {...p} />;
    case "textImport":
      return <TextImportDialog {...p} />;
    case "removeDuplicates":
      return <RemoveDuplicatesDialog {...p} />;
    case "textToColumns":
      return <TextToColumnsDialog {...p} />;
    case "shapeText":
      return <M.ShapeTextDialog {...p} />;
    case "pageSetup":
      return <PageSetupDialog {...p} />;
    case "shortcuts":
      return <ShortcutsDialog {...p} />;
    case "workbookStats":
      return <WorkbookStatsDialog {...p} />;
    case "createPivot":
      return <CreatePivotDialog {...p} />;
    case "pivotValueField":
      return <ValueFieldDialog {...p} />;
    case "pivotMove":
      return <MovePivotDialog {...p} />;
    case "pivotSource":
      return <ChangePivotSourceDialog {...p} />;
    case "pivotOptions":
      return <PivotOptionsDialog {...p} />;
    case "evaluateFormula":
      return <EvaluateFormulaDialog {...p} />;
    case "errorCheck":
      return <ErrorCheckDialog {...p} />;
    case "errorOptions":
      return <ErrorOptionsDialog {...p} />;
    case "sparkline":
      return <SparklineDialog {...p} />;
    case "sparklineItem":
      return <SparklineItemDialog {...p} />;
    case "sparklineHidden":
      return <SparklineHiddenDialog {...p} />;
    case "sparklineAxis":
      return <SparklineAxisDialog {...p} />;
  }
  return <M.AlertDialog {...p} arg={{ message: `"${state.id}" iletişim kutusu bulunamadı.` }} />;
}
