// Düzen iletişim kutuları: Satır Numaraları seçenekleri (imleçteki bölüme uygulanır).
import { useState } from "react";
import { Dialog, DialogButton, Field, FieldSet } from "@officehere/ui";
import { useBelge } from "../state/context";
import { applyPageSetup, sectionPageAt } from "../state/pageSetup";
import type { LineNumbering } from "../model/document";
import { Check } from "./FontDialog";

const TW_PER_CM = 1440 / 2.54;

export function LineNumbersDialog({ onClose }: { onClose: () => void }) {
  const { mainEditor, settings, updateSettings, fmt } = useBelge();
  const cur = sectionPageAt(mainEditor, settings.page, fmt.section).lineNumbers ?? null;
  const [on, setOn] = useState(!!cur);
  const [start, setStart] = useState(String(cur?.start ?? 1));
  const [dist, setDist] = useState(cur?.distance ? (cur.distance / TW_PER_CM).toFixed(2).replace(".", ",") : "Otomatik");
  const [by, setBy] = useState(String(cur?.countBy ?? 1));
  const [restart, setRestart] = useState<LineNumbering["restart"]>(cur?.restart ?? "newPage");
  const apply = () => {
    const cm = parseFloat(dist.replace(",", "."));
    const ln: LineNumbering | null = on
      ? { restart, start: Math.max(1, Math.round(Number(start) || 1)), countBy: Math.max(1, Math.round(Number(by) || 1)), ...(Number.isFinite(cm) && cm > 0 ? { distance: Math.round(cm * TW_PER_CM) } : {}) }
      : null;
    applyPageSetup(mainEditor, settings, updateSettings, (p) => ({ ...p, lineNumbers: ln }), "auto", fmt.section);
    onClose();
  };
  const radio = (v: LineNumbering["restart"], label: string) => (
    <label className="bl-radio">
      <input type="radio" name="lnrestart" checked={restart === v} disabled={!on} onChange={() => setRestart(v)} /> {label}
    </label>
  );
  return (
    <Dialog
      open
      title="Satır Numaraları"
      width={340}
      onClose={onClose}
      onSubmit={apply}
      footer={
        <>
          <DialogButton primary onClick={apply}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <Check label="Satır numaralandırması ekle" v={on} set={setOn} />
      <div className="bl-dlg-grid">
        <Field label="Başlangıç:">
          <input type="number" min={1} value={start} disabled={!on} onChange={(e) => setStart(e.target.value)} />
        </Field>
        <Field label="Metinden uzaklık:">
          <input value={dist} disabled={!on} onChange={(e) => setDist(e.target.value)} aria-label="Metinden uzaklık (cm)" />
        </Field>
        <Field label="Aralık:">
          <input type="number" min={1} value={by} disabled={!on} onChange={(e) => setBy(e.target.value)} />
        </Field>
      </div>
      <FieldSet legend="Numaralandırma">
        {radio("newPage", "Her sayfada yeniden başlat")}
        {radio("newSection", "Her bölümde yeniden başlat")}
        {radio("continuous", "Sürekli")}
      </FieldSet>
    </Dialog>
  );
}
