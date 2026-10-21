// Kurtarılan taslaklar bildirimi: kapanmış (ya da çökmüş) pencerelerden kalan kaydedilmemiş belgeler.
// Tek taslakta kısa çubuk; birden çok taslakta her biri ayrı satırda Geri Yükle / Sil ile listelenir.
// Kapat (×) çubuktaki taslakları siler: kapatılan taslak sonraki açılışlarda yeniden önerilmez.
import { MessageBar } from "@officehere/ui";
import { draftPreview, type StoredDraft } from "../state/drafts";

const when = (t: number) => new Date(t).toLocaleString("tr-TR");

interface Props {
  drafts: StoredDraft[];
  onRestore: (d: StoredDraft) => void;
  onDiscard: (d: StoredDraft) => void;
  onClose: () => void;
}

export function DraftBar({ drafts, onRestore, onDiscard, onClose }: Props) {
  if (!drafts.length) return null;
  if (drafts.length === 1) {
    const d = drafts[0];
    return (
      <MessageBar
        kind="warning"
        onClose={onClose}
        closeLabel="Taslağı sil ve kapat"
        actions={
          <>
            <button type="button" className="oh-btn is-primary" onClick={() => onRestore(d)}>
              Geri Yükle
            </button>
            <button type="button" className="oh-btn" onClick={() => onDiscard(d)}>
              Sil
            </button>
          </>
        }
      >
        <strong>Kurtarılan belge</strong> “{d.fileName}” için kaydedilmemiş değişiklikler bulundu ({when(d.savedAt)}).
      </MessageBar>
    );
  }
  return (
    <MessageBar kind="warning" onClose={onClose} closeLabel="Taslakları sil ve kapat">
      <strong>Belge Kurtarma</strong> — kaydedilmemiş değişiklikleri olan {drafts.length} belge bulundu:
      <ul className="bl-draftlist" aria-label="Kurtarılan belgeler">
        {drafts.map((d) => (
          <li key={d.key} className="bl-draftitem">
            <span className="bl-draftname">“{d.fileName}”</span>
            <span className="bl-draftprev">{draftPreview(d)}</span>
            <span className="bl-draftwhen">{when(d.savedAt)}</span>
            <button type="button" className="oh-btn is-primary" onClick={() => onRestore(d)} aria-label={`${d.fileName} belgesini geri yükle`}>
              Geri Yükle
            </button>
            <button type="button" className="oh-btn" onClick={() => onDiscard(d)} aria-label={`${d.fileName} taslağını sil`}>
              Sil
            </button>
          </li>
        ))}
      </ul>
    </MessageBar>
  );
}
