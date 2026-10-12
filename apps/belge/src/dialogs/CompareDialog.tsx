// Karşılaştır: özgün ve düzeltilmiş belgeyi seçip farkları izlenen değişiklik olarak yeni pencerede açar.
import { useState } from "react";
import { Dialog, DialogButton, Field, FieldSet, getUserName } from "@officehere/ui";
import { useBelge } from "../state/context";
import { openDocument } from "../io";
import type { BelgeDocument } from "../model/document";
import { compareDocuments } from "../editor/compare";
import { openInNewWindow } from "../mailmerge/data";

type Source = { kind: "current" } | { kind: "file"; file: File | null };

export function CompareDialog({ onClose }: { onClose: () => void }) {
  const { mainEditor, settings, notify } = useBelge();
  const [orig, setOrig] = useState<Source>({ kind: "current" });
  const [rev, setRev] = useState<Source>({ kind: "file", file: null });
  const [author, setAuthor] = useState(getUserName() || "Yazar");
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState("");

  const load = async (s: Source): Promise<BelgeDocument | null> => {
    if (s.kind === "current") return { content: mainEditor.getJSON(), settings };
    if (!s.file) return null;
    const { doc } = await openDocument(s.file.name, await s.file.arrayBuffer());
    return doc;
  };

  const run = async () => {
    if (orig.kind === "current" && rev.kind === "current") return setMsg("Aynı belgeyi kendisiyle karşılaştıramazsınız.");
    setBusy(true);
    try {
      const a = await load(orig);
      const b = await load(rev);
      if (!a || !b) return setMsg("Karşılaştırılacak iki belgeyi de seçin.");
      const content = compareDocuments(a.content, b.content, author.trim() || "Yazar");
      if (!openInNewWindow({ content, settings: { ...b.settings, mailMerge: null, protection: null } }, "Karşılaştırma Sonucu1")) notify("Karşılaştırma sonucu yeni pencerede açılamadı.", "error");
      else notify("Karşılaştırma sonucu yeni pencerede açıldı.", "success");
      onClose();
    } catch (err) {
      setMsg(`Belge okunamadı: ${(err as Error).message}`);
    } finally {
      setBusy(false);
    }
  };

  const picker = (label: string, s: Source, set: (s: Source) => void, name: string) => (
    <FieldSet legend={label}>
      <label className="bl-radio">
        <input type="radio" name={name} checked={s.kind === "current"} onChange={() => set({ kind: "current" })} /> Geçerli belge
      </label>
      <label className="bl-radio">
        <input type="radio" name={name} checked={s.kind === "file"} onChange={() => set({ kind: "file", file: s.kind === "file" ? s.file : null })} /> Dosyadan:
        <input
          type="file"
          accept=".docx,.docm,.dotx,.html,.htm,.txt"
          aria-label={`${label} dosyası`}
          // Dosya seçmek "Dosyadan" seçeneğini de seçer.
          onChange={(e) => set({ kind: "file", file: e.target.files?.[0] ?? null })}
        />
      </label>
    </FieldSet>
  );

  return (
    <Dialog
      open
      title="Belgeleri Karşılaştır"
      width={520}
      onClose={onClose}
      footer={
        <>
          <DialogButton primary disabled={busy} onClick={() => void run()}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      {picker("Özgün belge", orig, setOrig, "cmp-orig")}
      {picker("Düzeltilmiş belge", rev, setRev, "cmp-rev")}
      <Field label="Değişiklikleri şu adla etiketle:">
        <input value={author} onChange={(e) => setAuthor(e.target.value)} aria-label="Değişiklik yazarı" />
      </Field>
      <p className="bl-dlg-note">Sonuç, düzeltilmiş belgeye göre farkları izlenen değişiklik olarak gösteren yeni bir belgedir.</p>
      {msg && <div className="bl-dlg-msg is-error">{msg}</div>}
    </Dialog>
  );
}
