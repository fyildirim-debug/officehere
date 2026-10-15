// 4. turda eklenen iletişim kutularının yönlendiricisi (DialogHost'un devamı).
import type { DialogState } from "../state/context";
import { LineNumbersDialog } from "./LayoutDialogs";
import { EquationDialog } from "./EquationDialog";
import { ChartDialog } from "./ChartDialog";
import { CropDialog } from "./CropDialog";
import { ProtectDialog } from "./ProtectDialog";
import { CompareDialog } from "./CompareDialog";
import { InsertIndexDialog, MarkEntryDialog } from "./IndexDialogs";
import { NumberingValueDialog } from "./NumberingValueDialog";
import { DropCapDialog } from "./DropCapDialog";
import { DenetimPaneli } from "../denetim/DenetimPaneli";

export function ExtraDialogs({ dialog, onClose }: { dialog: DialogState; onClose: () => void }) {
  switch (dialog.kind) {
    case "lineNumbers":
      return <LineNumbersDialog onClose={onClose} />;
    case "equation":
      return <EquationDialog onClose={onClose} />;
    case "chart":
      return <ChartDialog onClose={onClose} />;
    case "crop":
      return <CropDialog onClose={onClose} />;
    case "protect":
      return <ProtectDialog onClose={onClose} />;
    case "compare":
      return <CompareDialog onClose={onClose} />;
    case "markEntry":
      return <MarkEntryDialog onClose={onClose} />;
    case "insertIndex":
      return <InsertIndexDialog onClose={onClose} />;
    case "numberingValue":
      return <NumberingValueDialog onClose={onClose} />;
    case "dropCap":
      return <DropCapDialog onClose={onClose} />;
    case "resmiDenetim":
      return <DenetimPaneli onClose={onClose} />;
    default:
      return null;
  }
}
