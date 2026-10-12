// Alıntılar ve Kaynakça: Kaynak Oluştur / Düzenle ve Kaynakları Yönet iletişim kutuları.
import { useState } from "react";
import { Dialog, DialogButton, Field } from "@officehere/ui";
import { useBelge } from "../state/context";
import { SOURCE_TYPE_LABEL, bibEntry, sourceTag, type Source, type SourceType } from "../model/sources";

function emptySource(): Source {
  return { tag: "", type: "book", author: "", title: "", year: "", publisher: "" };
}

/** Kaynak Oluştur (Alıntı Ekle › Yeni Kaynak Ekle) ve Kaynağı Düzenle. `insert` ise alıntı da eklenir. */
export function SourceDialog({ tag, insert, onClose }: { tag?: string; insert?: boolean; onClose: () => void }) {
  const ctx = useBelge();
  const sources = ctx.settings.sources ?? [];
  const editing = tag ? sources.find((s) => s.tag === tag) : undefined;
  const [s, setS] = useState<Source>(() => (editing ? { ...editing } : emptySource()));
  const set = (patch: Partial<Source>) => setS((x) => ({ ...x, ...patch }));
  const valid = !!s.title.trim() && (!!s.author.trim() || s.type === "web");
  const ok = () => {
    if (!valid) return;
    const others = sources.filter((x) => x.tag !== editing?.tag);
    const src: Source = {
      ...s,
      tag:
        editing?.tag ??
        sourceTag(
          s,
          others.map((x) => x.tag),
        ),
    };
    ctx.updateSettings((st) => ({ ...st, sources: [...(st.sources ?? []).filter((x) => x.tag !== src.tag), src] }));
    if (insert)
      ctx.mainEditor
        .chain()
        .focus()
        .insertContent({ type: "citation", attrs: { tag: src.tag } })
        .run();
    onClose();
  };
  const pubLabel = s.type === "article" ? "Dergi adı:" : s.type === "web" ? "Web sitesinin adı:" : s.type === "report" ? "Kurum:" : "Yayımcı:";
  return (
    <Dialog
      open
      title={editing ? "Kaynağı Düzenle" : "Kaynak Oluştur"}
      width={520}
      onClose={onClose}
      onSubmit={ok}
      footer={
        <>
          <DialogButton primary onClick={ok} disabled={!valid}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <Field label="Kaynak türü:" inline>
        <select aria-label="Kaynak türü" value={s.type} onChange={(e) => set({ type: e.target.value as SourceType })}>
          {(Object.keys(SOURCE_TYPE_LABEL) as SourceType[]).map((t) => (
            <option key={t} value={t}>
              {SOURCE_TYPE_LABEL[t]}
            </option>
          ))}
        </select>
      </Field>
      <div className="bl-src-grid">
        <label>
          Yazar:
          <input type="text" aria-label="Yazar" placeholder="Soyadı, Ad; Soyadı, Ad" value={s.author} onChange={(e) => set({ author: e.target.value })} />
        </label>
        <label>
          Başlık:
          <input type="text" aria-label="Başlık" value={s.title} onChange={(e) => set({ title: e.target.value })} />
        </label>
        <label>
          Yıl:
          <input type="text" aria-label="Yıl" value={s.year} onChange={(e) => set({ year: e.target.value })} />
        </label>
        <label>
          {pubLabel}
          <input type="text" aria-label="Yayımcı" value={s.publisher} onChange={(e) => set({ publisher: e.target.value })} />
        </label>
        {s.type === "book" && (
          <label>
            Şehir:
            <input type="text" aria-label="Şehir" value={s.city ?? ""} onChange={(e) => set({ city: e.target.value })} />
          </label>
        )}
        {s.type === "article" && (
          <>
            <label>
              Cilt:
              <input type="text" aria-label="Cilt" value={s.volume ?? ""} onChange={(e) => set({ volume: e.target.value })} />
            </label>
            <label>
              Sayı:
              <input type="text" aria-label="Sayı" value={s.issue ?? ""} onChange={(e) => set({ issue: e.target.value })} />
            </label>
            <label>
              Sayfalar:
              <input type="text" aria-label="Sayfalar" value={s.pages ?? ""} onChange={(e) => set({ pages: e.target.value })} />
            </label>
          </>
        )}
        {(s.type === "web" || s.type === "article" || s.type === "report") && (
          <label>
            URL:
            <input type="url" aria-label="URL" value={s.url ?? ""} onChange={(e) => set({ url: e.target.value })} />
          </label>
        )}
        {s.type === "web" && (
          <label>
            Erişim tarihi:
            <input type="text" aria-label="Erişim tarihi" placeholder="gg.aa.yyyy" value={s.accessed ?? ""} onChange={(e) => set({ accessed: e.target.value })} />
          </label>
        )}
      </div>
    </Dialog>
  );
}

/** Kaynakları Yönet: kaynak listesi, önizleme, Yeni / Düzenle / Sil, alıntı ekle. */
export function ManageSourcesDialog({ onClose }: { onClose: () => void }) {
  const ctx = useBelge();
  const sources = ctx.settings.sources ?? [];
  const style = ctx.settings.bibStyle ?? "apa";
  const [sel, setSel] = useState<string | null>(sources[0]?.tag ?? null);
  const cur = sources.find((s) => s.tag === sel);
  const entry = cur ? bibEntry(cur, style, sources.indexOf(cur)) : null;
  const remove = () => {
    if (!cur) return;
    ctx.updateSettings((st) => ({ ...st, sources: (st.sources ?? []).filter((x) => x.tag !== cur.tag) }));
    setSel(null);
  };
  return (
    <Dialog
      open
      title="Kaynak Yöneticisi"
      width={560}
      onClose={onClose}
      footer={
        <>
          <DialogButton onClick={() => ctx.openDialog("source")}>Yeni...</DialogButton>
          <DialogButton disabled={!cur} onClick={() => cur && ctx.openDialog("source", cur.tag)}>
            Düzenle...
          </DialogButton>
          <DialogButton disabled={!cur} onClick={remove}>
            Sil
          </DialogButton>
          <DialogButton
            disabled={!cur}
            onClick={() => {
              if (cur)
                ctx.mainEditor
                  .chain()
                  .focus()
                  .insertContent({ type: "citation", attrs: { tag: cur.tag } })
                  .run();
              onClose();
            }}
          >
            Alıntı Ekle
          </DialogButton>
          <DialogButton primary onClick={onClose}>
            Kapat
          </DialogButton>
        </>
      }
    >
      <div className="bl-xref-head">Geçerli liste ({sources.length})</div>
      <div className="bl-xref-list" role="listbox" aria-label="Kaynaklar">
        {sources.length === 0 && <div className="bl-xref-empty">Henüz kaynak yok. Yeni bir kaynak ekleyin.</div>}
        {sources.map((s) => (
          <div key={s.tag} role="option" aria-selected={s.tag === sel} className={s.tag === sel ? "is-selected" : undefined} onClick={() => setSel(s.tag)}>
            {s.author || "(yazar yok)"}; {s.title} ({s.year})
          </div>
        ))}
      </div>
      <div className="bl-xref-head">Önizleme ({style.toUpperCase()}):</div>
      <div className="bl-src-preview">
        {entry ? (
          <>
            {entry.before}
            <i>{entry.italic}</i>
            {entry.after}
          </>
        ) : (
          "—"
        )}
      </div>
    </Dialog>
  );
}
