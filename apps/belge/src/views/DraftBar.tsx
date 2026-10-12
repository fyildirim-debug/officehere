// Kurtarılan taslak bildirimi (Geri Yükle / Sil).
import { MessageBar } from "@officehere/ui";
import type { Draft } from "../state/useFileOps";

export function DraftBar({ draft, onRestore, onDiscard }: { draft: Draft; onRestore: () => void; onDiscard: () => void }) {
  return (
    <MessageBar
      kind="warning"
      onClose={onDiscard}
      actions={
        <>
          <button type="button" className="oh-btn is-primary" onClick={onRestore}>
            Geri Yükle
          </button>
          <button type="button" className="oh-btn" onClick={onDiscard}>
            Sil
          </button>
        </>
      }
    >
      <strong>Kurtarılan belge</strong> “{draft.fileName}” için kaydedilmemiş değişiklikler bulundu ({new Date(draft.savedAt).toLocaleString("tr-TR")}).
    </MessageBar>
  );
}
