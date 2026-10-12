// Açık iletişim kutusunu çizer.
import type { DialogState } from "../state/context";
import { AutoCorrectDialog } from "./AutoCorrectDialog";
import { StyleDialog } from "./StyleDialog";
import { TabsDialog } from "./TabsDialog";
import { CaptionDialog, CrossRefDialog, TableOfFiguresDialog } from "./ReferenceDialogs";
import { ManageSourcesDialog, SourceDialog } from "./SourceDialogs";
import { MergeDialog, RecipientsDialog } from "./MergeDialogs";
import { CellMarginsDialog, SortHost, TableSortDialog, TableToTextDialog, TextToTableDialog } from "./TableDialogs";
import { FontDialog } from "./FontDialog";
import { ParagraphDialog } from "./ParagraphDialog";
import { FindReplaceDialog } from "./FindReplaceDialog";
import { ColumnsDialog, PageSetupDialog } from "./PageSetupDialog";
import { BookmarkDialog, DateTimeDialog, InsertTableDialog, LinkDialog, SymbolDialog } from "./InsertDialogs";
import { AltTextDialog, ImageSizeDialog, PageBordersDialog, ShortcutsDialog, SplitCellsDialog, TablePropsDialog, WatermarkDialog, WordCountDialog, ZoomDialog, type ZoomTarget } from "./MiscDialogs";

interface Props {
  dialog: DialogState | null;
  onClose: () => void;
  zoom: number;
  onZoom: (z: ZoomTarget) => void;
  pages: number;
}

export function DialogHost({ dialog, onClose, zoom, onZoom, pages }: Props) {
  if (!dialog) return null;
  switch (dialog.kind) {
    case "font":
      return <FontDialog onClose={onClose} />;
    case "paragraph":
      return <ParagraphDialog onClose={onClose} />;
    case "findReplace":
    case "goto":
      return <FindReplaceDialog initialTab={dialog.kind === "goto" ? "goto" : (dialog.arg ?? "find")} onClose={onClose} />;
    case "pageSetup":
      return <PageSetupDialog initialTab={dialog.arg} onClose={onClose} />;
    case "columns":
      return <ColumnsDialog onClose={onClose} />;
    case "link":
      return <LinkDialog onClose={onClose} />;
    case "insertTable":
      return <InsertTableDialog onClose={onClose} />;
    case "symbol":
      return <SymbolDialog onClose={onClose} />;
    case "dateTime":
      return <DateTimeDialog onClose={onClose} />;
    case "bookmark":
      return <BookmarkDialog onClose={onClose} />;
    case "zoom":
      return <ZoomDialog current={zoom} onApply={onZoom} onClose={onClose} />;
    case "wordCount":
      return <WordCountDialog pages={pages} onClose={onClose} />;
    case "shortcuts":
      return <ShortcutsDialog onClose={onClose} />;
    case "watermark":
      return <WatermarkDialog onClose={onClose} />;
    case "pageBorders":
      return <PageBordersDialog onClose={onClose} />;
    case "splitCells":
      return <SplitCellsDialog onClose={onClose} />;
    case "tableProps":
      return <TablePropsDialog onClose={onClose} />;
    case "sort":
      return <SortHost onClose={onClose} />;
    case "tableSort":
      return <TableSortDialog onClose={onClose} />;
    case "textToTable":
      return <TextToTableDialog onClose={onClose} />;
    case "tableToText":
      return <TableToTextDialog onClose={onClose} />;
    case "cellMargins":
      return <CellMarginsDialog onClose={onClose} />;
    case "altText":
      return <AltTextDialog onClose={onClose} />;
    case "imageSize":
      return <ImageSizeDialog onClose={onClose} />;
    case "autocorrect":
      return <AutoCorrectDialog onClose={onClose} />;
    case "tabs":
      return <TabsDialog onClose={onClose} />;
    case "caption":
      return <CaptionDialog onClose={onClose} />;
    case "tableOfFigures":
      return <TableOfFiguresDialog onClose={onClose} />;
    case "crossRef":
      return <CrossRefDialog onClose={onClose} />;
    case "source":
      return <SourceDialog tag={dialog.arg && dialog.arg !== "insert" ? dialog.arg : undefined} insert={dialog.arg === "insert"} onClose={onClose} />;
    case "sources":
      return <ManageSourcesDialog onClose={onClose} />;
    case "recipients":
      return <RecipientsDialog onClose={onClose} />;
    case "merge":
      return <MergeDialog mode={dialog.arg === "print" ? "print" : "edit"} onClose={onClose} />;
    case "style":
      return <StyleDialog styleId={dialog.arg} onClose={onClose} />;
    default:
      return null;
  }
}
