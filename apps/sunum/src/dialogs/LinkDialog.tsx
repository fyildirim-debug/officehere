/** Köprü Ekle / Köprüyü Düzenle iletişim kutusu (Ekle > Bağlantı, Ctrl+K). */
import { useState } from "react";
import { Globe20Regular, Mail20Regular, SlideText20Regular } from "@fluentui/react-icons";
import { Dialog, DialogButton, Field } from "@officehere/ui";
import { JUMP_LABELS, normalizeUrl, parseLink } from "../model/links";
import type { Presentation } from "../model/types";
import { slideTitle } from "../show/SlideShow";
import type { LinkContext } from "../state/linkActions";

type Kind = "web" | "place" | "mail";

function initialState(c: LinkContext | null) {
  const link = c?.link;
  if (!link) return { kind: "web" as Kind, url: "", place: "", mail: "", subject: "" };
  const t = parseLink(link);
  if (t.kind === "slide") return { kind: "place" as Kind, url: "", place: link, mail: "", subject: "" };
  if (t.kind === "jump") return { kind: "place" as Kind, url: "", place: link, mail: "", subject: "" };
  if (t.url.startsWith("mailto:")) {
    const [addr, q] = t.url.slice(7).split("?");
    const subject = new URLSearchParams(q ?? "").get("subject") ?? "";
    return { kind: "mail" as Kind, url: "", place: "", mail: decodeURIComponent(addr), subject };
  }
  return { kind: "web" as Kind, url: t.url, place: "", mail: "", subject: "" };
}

export function LinkDialog({ context, pres, onApply, onClose }: { context: LinkContext | null; pres: Presentation; onApply: (c: LinkContext, link: string | null, text: string) => void; onClose: () => void }) {
  const open = !!context;
  const [st, setSt] = useState(() => initialState(context));
  const [text, setText] = useState(context?.text ?? "");
  const [last, setLast] = useState<LinkContext | null>(null);
  if (context !== last) {
    setLast(context);
    if (context) {
      setSt(initialState(context));
      setText(context.text);
    }
  }
  if (!context) return null;
  const editing = !!context.link;

  const address = (): string => {
    if (st.kind === "web") return normalizeUrl(st.url);
    if (st.kind === "place") return st.place;
    const m = st.mail.trim();
    if (!m) return "";
    return `mailto:${m}${st.subject ? `?subject=${encodeURIComponent(st.subject)}` : ""}`;
  };
  const ok = () => {
    const link = address();
    if (!link) {
      onClose();
      return;
    }
    onApply(context, link, context.textEditable ? text : context.text);
    onClose();
  };

  const places: { id: string; label: string; indent?: boolean }[] = [
    { id: "#first", label: JUMP_LABELS.first },
    { id: "#last", label: JUMP_LABELS.last },
    { id: "#next", label: JUMP_LABELS.next },
    { id: "#prev", label: JUMP_LABELS.prev },
    ...pres.slides.map((s, i) => ({ id: `#slide:${s.id}`, label: `${i + 1}. ${slideTitle(s) || `Slayt ${i + 1}`}`, indent: true })),
  ];

  return (
    <Dialog
      open={open}
      title={editing ? "Köprüyü Düzenle" : "Köprü Ekle"}
      onClose={onClose}
      onSubmit={ok}
      width={600}
      footer={
        <>
          {editing && (
            <DialogButton
              onClick={() => {
                onApply(context, null, context.text);
                onClose();
              }}
            >
              Bağlantıyı Kaldır
            </DialogButton>
          )}
          <DialogButton primary onClick={ok}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <Field label="Görüntülenecek metin:">
        <input type="text" aria-label="Görüntülenecek metin" value={context.textEditable ? text : context.text} disabled={!context.textEditable} onChange={(e) => setText(e.target.value)} style={{ width: "100%", height: 26 }} />
      </Field>
      <div className="link-dlg">
        <div className="link-dlg-nav" role="tablist" aria-label="Bağlantı ver">
          <div className="link-dlg-cap">Bağlantı ver:</div>
          {(
            [
              ["web", "Var Olan Dosya veya Web Sayfası", <Globe20Regular key="w" />],
              ["place", "Bu Belgede Yer", <SlideText20Regular key="p" />],
              ["mail", "E-posta Adresi", <Mail20Regular key="m" />],
            ] as [Kind, string, React.ReactNode][]
          ).map(([k, label, icon]) => (
            <button key={k} type="button" role="tab" aria-selected={st.kind === k} className={`link-dlg-tab${st.kind === k ? " is-active" : ""}`} onClick={() => setSt({ ...st, kind: k })}>
              {icon}
              <span>{label}</span>
            </button>
          ))}
        </div>
        <div className="link-dlg-body">
          {st.kind === "web" && (
            <Field label="Adres:">
              <input type="text" aria-label="Adres" placeholder="https://" value={st.url} onChange={(e) => setSt({ ...st, url: e.target.value })} style={{ width: "100%", height: 26 }} autoFocus />
            </Field>
          )}
          {st.kind === "place" && (
            <>
              <div className="link-dlg-cap">Belgede bir yer seçin:</div>
              <div className="link-dlg-list" role="listbox" aria-label="Belgede bir yer seçin">
                {places.map((p) => (
                  <div key={p.id} role="option" aria-selected={st.place === p.id} className={`list-row${st.place === p.id ? " is-selected" : ""}`} style={{ paddingLeft: p.indent ? 22 : 6 }} onClick={() => setSt({ ...st, place: p.id })} onDoubleClick={() => setTimeout(ok)}>
                    {p.label}
                  </div>
                ))}
              </div>
            </>
          )}
          {st.kind === "mail" && (
            <>
              <Field label="E-posta adresi:">
                <input type="text" aria-label="E-posta adresi" value={st.mail} onChange={(e) => setSt({ ...st, mail: e.target.value })} style={{ width: "100%", height: 26 }} />
              </Field>
              <Field label="Konu:">
                <input type="text" aria-label="Konu" value={st.subject} onChange={(e) => setSt({ ...st, subject: e.target.value })} style={{ width: "100%", height: 26 }} />
              </Field>
            </>
          )}
        </div>
      </div>
    </Dialog>
  );
}
