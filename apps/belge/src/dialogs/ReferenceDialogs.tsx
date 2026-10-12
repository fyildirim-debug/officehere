// Başvurular iletişim kutuları: Resim Yazısı, Şekiller Tablosu, Çapraz Başvuru.
import { useMemo, useState } from "react";
import { Dialog, DialogButton, Field, FieldSet } from "@officehere/ui";
import { NodeSelection } from "@tiptap/pm/state";
import { useBelge } from "../state/context";
import { CAPTION_LABELS, collectCaptions, insertCaption, insertCrossRef, refTargets } from "../editor/references";
import { CROSSREF_KIND_LABEL, type CrossRefKind } from "../editor/fields";
import { Check } from "./FontDialog";

const okCancel = (ok: () => void, close: () => void, okLabel = "Tamam", disabled = false) => (
  <>
    <DialogButton primary onClick={ok} disabled={disabled}>
      {okLabel}
    </DialogButton>
    <DialogButton onClick={close}>{okLabel === "Ekle" ? "Kapat" : "İptal"}</DialogButton>
  </>
);

/** Kullanıcının eklediği etiketler bu oturumda listede kalır. */
const extraLabels: string[] = [];

export function CaptionDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const sel = editor.state.selection;
  const inTable = editor.isActive("table");
  const isImage = sel instanceof NodeSelection && sel.node.type.name === "image";
  const [label, setLabel] = useState(inTable ? "Tablo" : "Şekil");
  const [text, setText] = useState("");
  const [position, setPosition] = useState<"above" | "below">(inTable ? "above" : "below");
  const [exclude, setExclude] = useState(false);
  const [newLabel, setNewLabel] = useState<string | null>(null);
  const labels = [...CAPTION_LABELS, ...extraLabels];
  const next = collectCaptions(editor.state.doc, label).length + 1;
  const ok = () => {
    insertCaption(editor, label, text.trim(), position, exclude);
    onClose();
  };
  return (
    <Dialog open title="Resim Yazısı" width={440} onClose={onClose} onSubmit={ok} footer={okCancel(ok, onClose)}>
      <Field label="Resim yazısı:">
        <div className="bl-cap-row">
          <span className="bl-cap-prefix">{exclude ? next : `${label} ${next}`}</span>
          <input type="text" aria-label="Resim yazısı" value={text} onChange={(e) => setText(e.target.value)} autoFocus />
        </div>
      </Field>
      <FieldSet legend="Seçenekler">
        <div className="bl-dlg-grid">
          <Field label="Etiket:">
            <select aria-label="Etiket" value={label} onChange={(e) => setLabel(e.target.value)}>
              {labels.map((l) => (
                <option key={l} value={l}>
                  {l}
                </option>
              ))}
            </select>
          </Field>
          <Field label="Konum:">
            <select aria-label="Konum" value={position} disabled={!inTable && !isImage} onChange={(e) => setPosition(e.target.value as "above" | "below")}>
              <option value="above">Seçili öğenin üstünde</option>
              <option value="below">Seçili öğenin altında</option>
            </select>
          </Field>
        </div>
        <Check label="Etiketi resim yazısından çıkar" v={exclude} set={setExclude} />
        {newLabel === null ? (
          <DialogButton onClick={() => setNewLabel("")}>Yeni Etiket...</DialogButton>
        ) : (
          <div className="bl-cap-row">
            <input type="text" aria-label="Yeni etiket" value={newLabel} onChange={(e) => setNewLabel(e.target.value)} />
            <DialogButton
              disabled={!newLabel.trim()}
              onClick={() => {
                const l = newLabel.trim();
                if (!labels.includes(l)) extraLabels.push(l);
                setLabel(l);
                setNewLabel(null);
              }}
            >
              Ekle
            </DialogButton>
          </div>
        )}
      </FieldSet>
    </Dialog>
  );
}

export function TableOfFiguresDialog({ onClose }: { onClose: () => void }) {
  const ctx = useBelge();
  const { editor } = ctx;
  const present = useMemo(() => [...new Set(collectCaptions(editor.state.doc).map((c) => c.label))], [editor]);
  const [label, setLabel] = useState(present[0] ?? "Şekil");
  const [pages, setPages] = useState(true);
  const preview = collectCaptions(editor.state.doc, label).slice(0, 4);
  const ok = () => {
    editor
      .chain()
      .focus()
      .insertContent({ type: "toc", attrs: { kind: "tof", label, title: "", entries: [], showPages: pages } })
      .run();
    requestAnimationFrame(() => requestAnimationFrame(ctx.updateToc));
    onClose();
  };
  return (
    <Dialog open title="Şekiller Tablosu" width={460} onClose={onClose} onSubmit={ok} footer={okCancel(ok, onClose)}>
      <FieldSet legend="Baskı Önizleme">
        <div className="bl-tof-preview">
          {(preview.length ? preview.map((c) => c.text) : [`${label} 1: Metin`, `${label} 2: Metin`]).map((t, i) => (
            <div key={i}>
              <span>{t}</span>
              <span className="bl-tof-dots" />
              {pages && <span>{i + 1}</span>}
            </div>
          ))}
        </div>
      </FieldSet>
      <Check label="Sayfa numaralarını göster" v={pages} set={setPages} />
      <Field label="Resim yazısı etiketi:" inline>
        <select aria-label="Resim yazısı etiketi" value={label} onChange={(e) => setLabel(e.target.value)}>
          {[...new Set([...CAPTION_LABELS, ...present, ...extraLabels])].map((l) => (
            <option key={l} value={l}>
              {l}
            </option>
          ))}
        </select>
      </Field>
    </Dialog>
  );
}

const REF_TYPES = ["Şekil", "Tablo", "Denklem", "Başlık", "Yer İşareti"];

export function CrossRefDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const [type, setType] = useState("Şekil");
  const [kind, setKind] = useState<CrossRefKind>("label");
  const [link, setLink] = useState(true);
  const [index, setIndex] = useState(0);
  const targets = useMemo(() => refTargets(editor.state.doc, type), [editor, type]);
  const kinds: CrossRefKind[] = type === "Başlık" || type === "Yer İşareti" ? ["full", "page", "position"] : ["full", "label", "text", "page", "position"];
  const kindLabel = (k: CrossRefKind) =>
    type === "Başlık" && k === "full" ? "Başlık metni" : type === "Yer İşareti" && k === "full" ? "Yer işareti metni" : k === "full" ? "Tüm resim yazısı" : CROSSREF_KIND_LABEL[k];
  const insert = () => {
    const t = targets[index];
    if (!t) return;
    insertCrossRef(editor, t, kinds.includes(kind) ? kind : "full", link);
  };
  return (
    <Dialog open title="Çapraz Başvuru" width={480} onClose={onClose} onSubmit={insert} footer={okCancel(insert, onClose, "Ekle", !targets.length)}>
      <div className="bl-dlg-grid">
        <Field label="Başvuru türü:">
          <select
            aria-label="Başvuru türü"
            value={type}
            onChange={(e) => {
              setType(e.target.value);
              setIndex(0);
              setKind(e.target.value === "Başlık" || e.target.value === "Yer İşareti" ? "full" : "label");
            }}
          >
            {REF_TYPES.map((t) => (
              <option key={t} value={t}>
                {t}
              </option>
            ))}
          </select>
        </Field>
        <Field label="Başvuruyu ekle:">
          <select aria-label="Başvuruyu ekle" value={kinds.includes(kind) ? kind : "full"} onChange={(e) => setKind(e.target.value as CrossRefKind)}>
            {kinds.map((k) => (
              <option key={k} value={k}>
                {kindLabel(k)}
              </option>
            ))}
          </select>
        </Field>
      </div>
      <Check label="Köprü olarak ekle" v={link} set={setLink} />
      <div className="bl-xref-head">Hangi {type.toLocaleLowerCase("tr-TR")} için:</div>
      <div className="bl-xref-list" role="listbox" aria-label="Başvuru hedefleri">
        {targets.length === 0 && <div className="bl-xref-empty">Belgede bu türde öğe yok.</div>}
        {targets.map((t, i) => (
          <div key={`${t.pos}-${i}`} role="option" aria-selected={i === index} className={i === index ? "is-selected" : undefined} onClick={() => setIndex(i)} onDoubleClick={insert}>
            {t.label}
          </div>
        ))}
      </div>
    </Dialog>
  );
}
