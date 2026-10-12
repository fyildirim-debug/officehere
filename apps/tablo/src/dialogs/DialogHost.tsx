/** İletişim kutusu yöneticisi: kimliğe göre ilgili iletişim kutusunu çizer. */
import { useEffect } from "react";
import type { Controller } from "../controller";
import type { FileActions } from "../FileView";
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

export interface DialogState {
  id: string;
  arg?: unknown;
}

/** Dialog açmadan doğrudan çalışan eylemler. */
const ACTIONS: Record<string, (ctl: Controller) => void> = {
  toggleFilter: (ctl) => D.toggleAutoFilter(ctl),
};

export function DialogHost({ ctl, state, onClose, files, onBackstage }: { ctl: Controller; state: DialogState | null; onClose: () => void; files: FileActions; onBackstage?: (page: string) => void }) {
  const id = state?.id;
  useEffect(() => {
    if (!id) return;
    if (ACTIONS[id]) {
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
    if (!state) return;
    const raf = requestAnimationFrame(() => {
      const el = document.querySelector<HTMLInputElement>(".oh-dialog [data-autofocus]");
      if (!el) return;
      el.focus();
      if (typeof el.select === "function") el.select();
    });
    return () => cancelAnimationFrame(raf);
  }, [state]);
  if (!state || ACTIONS[state.id] || state.id === "printPreview" || state.id === "about") return null;
  const p = { ctl, onClose, arg: state.arg };
  switch (state.id) {
    case "alert":
      return <M.AlertDialog {...p} />;
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
  }
  return <M.AlertDialog {...p} arg={{ message: `"${state.id}" iletişim kutusu bulunamadı.` }} />;
}
