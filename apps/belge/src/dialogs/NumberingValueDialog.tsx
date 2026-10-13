// Numaralandırma Değerini Ayarla: imleçteki öğeden yeni liste başlatır ya da önceki listeden devam eder.
import { useMemo, useState } from "react";
import { Dialog, DialogButton, Field, FieldSet } from "@officehere/ui";
import { useBelge } from "../state/context";
import { continueNumbering, itemNumber, orderedListAt, previousOrderedList, restartNumbering } from "../editor/listNumbering";

export function NumberingValueDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const at = useMemo(() => orderedListAt(editor.state), [editor]);
  const prev = useMemo(() => (at ? previousOrderedList(editor.state, at) : null), [editor, at]);
  const [mode, setMode] = useState<"new" | "continue">("new");
  const [value, setValue] = useState(String(at ? itemNumber(at) : 1));
  const continued = prev ? (Number(prev.node.attrs.start) || 1) + prev.node.childCount + (at?.index ?? 0) : null;
  const n = Math.round(Number(value));
  const valid = mode === "continue" ? !!prev : Number.isFinite(n) && n >= 0 && n <= 32767;
  const apply = () => {
    if (!at || !valid) return;
    if (mode === "continue") continueNumbering(editor);
    else restartNumbering(editor, n);
    onClose();
  };
  return (
    <Dialog
      open
      title="Numaralandırma Değerini Ayarla"
      width={360}
      onClose={onClose}
      onSubmit={apply}
      footer={
        <>
          <DialogButton primary onClick={apply} disabled={!at || !valid}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      {!at ? (
        <p className="bl-dlg-note">İmleç numaralı bir listede değil.</p>
      ) : (
        <>
          <FieldSet legend="Seçenek">
            <label className="bl-radio">
              <input type="radio" name="numval" checked={mode === "new"} onChange={() => setMode("new")} /> Yeni liste başlat
            </label>
            <label className="bl-radio">
              <input type="radio" name="numval" checked={mode === "continue"} disabled={!prev} onChange={() => setMode("continue")} /> Önceki listeden devam et
            </label>
          </FieldSet>
          <Field label="Değeri şuna ayarla:">
            <input
              type="number"
              min={0}
              max={32767}
              aria-label="Değeri şuna ayarla"
              value={mode === "continue" && continued !== null ? String(continued) : value}
              disabled={mode === "continue"}
              onChange={(e) => setValue(e.target.value)}
            />
          </Field>
        </>
      )}
    </Dialog>
  );
}
