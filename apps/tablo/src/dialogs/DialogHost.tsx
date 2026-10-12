import type { Controller } from "../controller";
import type { FileActions } from "../FileView";
export interface DialogState { id: string; arg?: unknown }
export function DialogHost({ ctl, state, onClose, files }: { ctl: Controller; state: DialogState | null; onClose: () => void; files: FileActions }) {
  void ctl; void state; void onClose; void files;
  return null;
}
