// Dizin iletişim kutuları: Dizin Girdisini İşaretle (açık kalır; İşaretle / Tümünü İşaretle) ve Dizin Ekle.
import { useState } from "react";
import { Dialog, DialogButton, Field } from "@officehere/ui";
import { useBelge } from "../state/context";
import { indexUpdate } from "../editor/indexing";

export function MarkEntryDialog({ onClose }: { onClose: () => void }) {
  const { mainEditor: editor, notify } = useBelge();
  const { from, to } = editor.state.selection;
  const [main, setMain] = useState(editor.state.doc.textBetween(from, to, " ").trim().slice(0, 120));
  const [sub, setSub] = useState("");
  const [msg, setMsg] = useState("");
  const mark = () => {
    if (!main.trim()) return setMsg("Ana girdiyi yazın.");
    editor.chain().markIndexEntry(main, sub).run();
    setMsg(`İşaretlendi: ${main}${sub ? ` › ${sub}` : ""}`);
  };
  // Tümünü İşaretle: ana girdi metninin belgedeki her geçişi işaretlenir (paragraf başına ilk geçiş).
  const markAll = () => {
    const word = main.trim();
    if (!word) return setMsg("Ana girdiyi yazın.");
    const lower = word.toLocaleLowerCase("tr-TR");
    const at: number[] = [];
    editor.state.doc.descendants((n, pos) => {
      if (!n.isTextblock) return true;
      const text = n.textContent;
      let k = text.toLocaleLowerCase("tr-TR").indexOf(lower);
      let off = 0;
      while (k >= 0) {
        // Metin konumunu belge konumuna çevir (satır içi düğümler atlanarak).
        let rem = k + lower.length;
        n.forEach((c, cOff) => {
          if (rem < 0) return;
          const len = c.isText ? (c.text?.length ?? 0) : 0;
          if (rem <= len && c.isText) {
            at.push(pos + 1 + cOff + rem);
            rem = -1;
          } else rem -= len;
        });
        off = k + lower.length;
        k = text.toLocaleLowerCase("tr-TR").indexOf(lower, off);
      }
      return false;
    });
    if (!at.length) return setMsg("Belgede bu metin bulunamadı.");
    const tr = editor.state.tr;
    for (const p of at.sort((a, b) => b - a)) tr.insert(p, editor.state.schema.nodes.indexEntry.create({ main: word, sub: sub.trim() || null }));
    editor.view.dispatch(tr);
    setMsg(`${at.length} girdi işaretlendi.`);
    notify(`"${word}" için ${at.length} dizin girdisi işaretlendi.`, "success");
  };
  return (
    <Dialog
      open
      title="Dizin Girdisini İşaretle"
      width={400}
      onClose={onClose}
      onSubmit={mark}
      footer={
        <>
          <DialogButton primary onClick={mark}>
            İşaretle
          </DialogButton>
          <DialogButton onClick={markAll}>Tümünü İşaretle</DialogButton>
          <DialogButton onClick={onClose}>Kapat</DialogButton>
        </>
      }
    >
      <Field label="Ana girdi:">
        <input value={main} autoFocus onChange={(e) => setMain(e.target.value)} aria-label="Ana girdi" />
      </Field>
      <Field label="Alt girdi:">
        <input value={sub} onChange={(e) => setSub(e.target.value)} aria-label="Alt girdi" />
      </Field>
      <p className="bl-dlg-note">Girdi alanları yalnız Tümünü Göster (Ctrl+*) açıkken görünür.</p>
      {msg && <div className="bl-dlg-msg">{msg}</div>}
    </Dialog>
  );
}

export function InsertIndexDialog({ onClose }: { onClose: () => void }) {
  const { mainEditor: editor, controller } = useBelge();
  const [cols, setCols] = useState(2);
  const apply = () => {
    editor.chain().focus().insertIndex(Math.max(1, Math.min(4, cols))).run();
    // Sayfa numaraları sayfalama sonrası hesaplanır.
    requestAnimationFrame(() =>
      requestAnimationFrame(() => {
        const tr = indexUpdate(editor.state, (p) => controller.pageOf(p));
        if (tr) editor.view.dispatch(tr);
      }),
    );
    onClose();
  };
  return (
    <Dialog
      open
      title="Dizin"
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
      <Field label="Sütunlar:">
        <input type="number" min={1} max={4} value={cols} onChange={(e) => setCols(Number(e.target.value) || 1)} aria-label="Sütunlar" />
      </Field>
      <Field label="Dil:">
        <select disabled value="tr">
          <option value="tr">Türkçe (Türkiye)</option>
        </select>
      </Field>
    </Dialog>
  );
}
