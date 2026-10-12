// Adres mektup birleştirme iletişim kutuları: Alıcı Listesini Düzenle ve Yeni Belgede Birleştir.
import { useState } from "react";
import { Dialog, DialogButton } from "@officehere/ui";
import { useBelge } from "../state/context";
import { useMerge } from "../mailmerge/useMailMerge";
import { mergeDocument, openInNewWindow } from "../mailmerge/data";

export function RecipientsDialog({ onClose }: { onClose: () => void }) {
  const m = useMerge();
  const [excluded, setExcluded] = useState<number[]>(m.excluded);
  const list = m.list;
  if (!list) return null;
  const toggle = (i: number) => setExcluded((x) => (x.includes(i) ? x.filter((k) => k !== i) : [...x, i]));
  const ok = () => {
    m.setExcluded(excluded);
    onClose();
  };
  return (
    <Dialog
      open
      title="Adres Mektup Birleştirme Alıcıları"
      width={640}
      onClose={onClose}
      onSubmit={ok}
      footer={
        <DialogButton primary onClick={ok}>
          Tamam
        </DialogButton>
      }
    >
      <p className="bl-merge-note">Bu, birleştirmede kullanılacak alıcı listesidir. Listeden çıkarmak istediğiniz alıcıların onay işaretini kaldırın.</p>
      <div className="bl-merge-table">
        <table>
          <thead>
            <tr>
              <th>
                <input type="checkbox" aria-label="Tümü" checked={!excluded.length} onChange={(e) => setExcluded(e.target.checked ? [] : list.records.map((_, i) => i))} />
              </th>
              {list.fields.map((f) => (
                <th key={f}>{f}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {list.records.map((r, i) => (
              <tr key={i}>
                <td>
                  <input type="checkbox" aria-label={`Kayıt ${i + 1}`} checked={!excluded.includes(i)} onChange={() => toggle(i)} />
                </td>
                {list.fields.map((f) => (
                  <td key={f}>{r[f]}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="bl-merge-note">Veri kaynağı: {list.name}</p>
    </Dialog>
  );
}

export function MergeDialog({ mode, onClose }: { mode: "edit" | "print"; onClose: () => void }) {
  const ctx = useBelge();
  const m = useMerge();
  const [which, setWhich] = useState<"all" | "current" | "range">("all");
  const [from, setFrom] = useState(1);
  const [to, setTo] = useState(m.list?.records.length ?? 1);
  const ok = () => {
    const all = m.list?.records ?? [];
    let recs = all.filter((_, i) => !m.excluded.includes(i));
    if (which === "current") recs = [all[m.index]].filter(Boolean);
    else if (which === "range") recs = all.slice(Math.max(0, from - 1), Math.max(0, to)).filter((_, k) => !m.excluded.includes(k + from - 1));
    if (!recs.length) return ctx.notify("Birleştirilecek kayıt yok.", "warning");
    const content = mergeDocument(ctx.mainEditor.getJSON(), recs);
    const opened = openInNewWindow({ content, settings: { ...ctx.settings, mailMerge: null } }, "Mektuplar1", mode === "print" ? "print" : undefined);
    if (!opened) ctx.notify("Birleştirilen belge açılamadı.", "error");
    else ctx.notify(`${recs.length} kayıt yeni belgede birleştirildi.`, "success");
    onClose();
  };
  const radio = (v: typeof which, label: string) => (
    <label className="bl-radio">
      <input type="radio" name="merge-which" checked={which === v} onChange={() => setWhich(v)} /> {label}
    </label>
  );
  return (
    <Dialog
      open
      title={mode === "print" ? "Yazıcıda Birleştir" : "Yeni Belgede Birleştir"}
      width={360}
      onClose={onClose}
      onSubmit={ok}
      footer={
        <>
          <DialogButton primary onClick={ok}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <div className="bl-merge-which">
        <span>Kayıtları birleştir</span>
        {radio("all", "Tümü")}
        {radio("current", "Geçerli kayıt")}
        <div className="bl-merge-range">
          {radio("range", "Başlangıç:")}
          <input type="number" aria-label="Başlangıç" min={1} value={from} onChange={(e) => setFrom(Number(e.target.value) || 1)} disabled={which !== "range"} />
          <span>Bitiş:</span>
          <input type="number" aria-label="Bitiş" min={1} value={to} onChange={(e) => setTo(Number(e.target.value) || 1)} disabled={which !== "range"} />
        </div>
      </div>
    </Dialog>
  );
}
