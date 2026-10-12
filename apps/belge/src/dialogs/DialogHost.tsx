// Açık iletişim kutusunu çizer.
import type { DialogState } from "../state/context";
import { FontDialog } from "./FontDialog";
import { ParagraphDialog } from "./ParagraphDialog";
import { FindReplaceDialog } from "./FindReplaceDialog";
import { ColumnsDialog, PageSetupDialog } from "./PageSetupDialog";
import { BookmarkDialog, DateTimeDialog, InsertTableDialog, LinkDialog, SymbolDialog } from "./InsertDialogs";
import { AltTextDialog, ImageSizeDialog, PageBordersDialog, ShortcutsDialog, SortDialog, TablePropsDialog, WatermarkDialog, WordCountDialog, ZoomDialog, type ZoomTarget } from "./MiscDialogs";

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
    case "tableProps":
      return <TablePropsDialog onClose={onClose} />;
    case "sort":
      return <SortDialog onClose={onClose} />;
    case "altText":
      return <AltTextDialog onClose={onClose} />;
    case "imageSize":
      return <ImageSizeDialog onClose={onClose} />;
    default:
      return null;
  }
}
