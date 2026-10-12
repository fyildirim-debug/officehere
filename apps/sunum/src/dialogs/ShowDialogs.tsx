/**
 * Slayt gösterisi iletişim kutuları: Gösteriyi Ayarla, Özel Gösteriler
 * (tanımla/düzenle), Zamanlama Provası sonucu ve slayt boyutu ölçekleme sorusu.
 */
import { useState } from "react";
import { Dialog, DialogButton, FieldSet } from "@officehere/ui";
import { uid } from "../model/factory";
import { resizePresentation } from "../model/presentation";
import type { CustomShow, Presentation } from "../model/types";
import { formatTime } from "../show/showKeys";
import { SHOW_INK_COLORS } from "../show/ShowInk";
import type { RehearsalResult } from "../show/Rehearsal";
import { uniqueShowName, type ShowSettings } from "../show/settings";
import { slideTitle } from "../show/SlideShow";
import type { Actions } from "../state/actions";

function Radio({ label, checked, onChange }: { label: string; checked: boolean; onChange: () => void }) {
  return (
    <label className="dlg-row">
      <input type="radio" checked={checked} onChange={onChange} /> {label}
    </label>
  );
}

function Check({ label, checked, onChange, disabled }: { label: string; checked: boolean; onChange: (v: boolean) => void; disabled?: boolean }) {
  return (
    <label className={`dlg-row${disabled ? " is-disabled" : ""}`}>
      <input type="checkbox" checked={checked} disabled={disabled} onChange={(e) => onChange(e.target.checked)} /> {label}
    </label>
  );
}

const LASER_COLORS = [
  { name: "Kırmızı", hex: "#FF0000" },
  { name: "Yeşil", hex: "#00B050" },
  { name: "Mavi", hex: "#0070C0" },
];

function ColorSelect({ label, value, options, onChange }: { label: string; value: string; options: { name: string; hex: string }[]; onChange: (v: string) => void }) {
  return (
    <label className="dlg-row">
      <span style={{ minWidth: 150 }}>{label}</span>
      <span className="show-swatch" style={{ background: value }} />
      <select aria-label={label} value={value} onChange={(e) => onChange(e.target.value)}>
        {options.map((c) => (
          <option key={c.hex} value={c.hex}>
            {c.name}
          </option>
        ))}
      </select>
    </label>
  );
}

export function ShowSetupDialog({ open, onClose, value, onChange, pres }: { open: boolean; onClose: () => void; value: ShowSettings; onChange: (v: ShowSettings) => void; pres: Presentation }) {
  const [v, setV] = useState(value);
  const [lastOpen, setLastOpen] = useState(false);
  const n = pres.slides.length;
  const shows = pres.customShows ?? [];
  if (open !== lastOpen) {
    setLastOpen(open);
    if (open) setV(value);
  }
  const ok = () => {
    let range = v.range;
    if (range.kind === "range") range = { kind: "range", from: Math.max(1, Math.min(n, range.from)), to: Math.max(1, Math.min(n, Math.max(range.from, range.to))) };
    if (range.kind === "custom" && !shows.some((c) => c.id === (range as { id: string }).id)) range = { kind: "all" };
    onChange({ ...v, range });
    onClose();
  };
  const kiosk = v.type === "kiosk";
  const from = v.range.kind === "range" ? v.range.from : 1;
  const to = v.range.kind === "range" ? v.range.to : n;
  const set = (p: Partial<ShowSettings>) => setV((x) => ({ ...x, ...p }));
  return (
    <Dialog
      open={open}
      title="Gösteriyi Ayarla"
      onClose={onClose}
      onSubmit={ok}
      width={640}
      footer={
        <>
          <DialogButton primary onClick={ok}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <div className="dlg-cols">
        <div>
          <FieldSet legend="Gösteri türü">
            <Radio label="Sunucu tarafından sunulan (tam ekran)" checked={v.type === "speaker"} onChange={() => set({ type: "speaker" })} />
            <Radio label="Birey tarafından göz atılan (pencere)" checked={v.type === "browsed"} onChange={() => set({ type: "browsed" })} />
            <Radio label="Bilgi noktasında göz atılan (tam ekran)" checked={kiosk} onChange={() => set({ type: "kiosk" })} />
          </FieldSet>
          <FieldSet legend="Gösteri seçenekleri">
            <Check label="'Esc' tuşuna basılana kadar sürekli döngü" checked={v.loop || kiosk} disabled={kiosk} onChange={(c) => set({ loop: c })} />
            <Check label="Anlatımsız göster" checked={v.noNarration} onChange={(c) => set({ noNarration: c })} />
            <Check label="Animasyonsuz göster" checked={v.noAnimations} onChange={(c) => set({ noAnimations: c })} />
            <ColorSelect label="Kalem rengi:" value={v.penColor} options={SHOW_INK_COLORS} onChange={(c) => set({ penColor: c })} />
            <ColorSelect label="Lazer işaretçisi rengi:" value={v.laserColor} options={LASER_COLORS} onChange={(c) => set({ laserColor: c })} />
          </FieldSet>
        </div>
        <div>
          <FieldSet legend="Slaytları göster">
            <Radio label="Tümü" checked={v.range.kind === "all"} onChange={() => set({ range: { kind: "all" } })} />
            <label className="dlg-row">
              <input type="radio" checked={v.range.kind === "range"} onChange={() => set({ range: { kind: "range", from, to } })} /> Başlangıç:
              <input type="number" aria-label="Başlangıç" min={1} max={n} value={from} style={{ width: 52 }} onChange={(e) => set({ range: { kind: "range", from: Number(e.target.value) || 1, to } })} />
              Bitiş:
              <input type="number" aria-label="Bitiş" min={1} max={n} value={to} style={{ width: 52 }} onChange={(e) => set({ range: { kind: "range", from, to: Number(e.target.value) || n } })} />
            </label>
            <label className={`dlg-row${shows.length ? "" : " is-disabled"}`}>
              <input type="radio" disabled={!shows.length} checked={v.range.kind === "custom"} onChange={() => shows[0] && set({ range: { kind: "custom", id: shows[0].id } })} /> Özel gösteri:
              <select aria-label="Özel gösteri" disabled={!shows.length} value={v.range.kind === "custom" ? v.range.id : (shows[0]?.id ?? "")} onChange={(e) => set({ range: { kind: "custom", id: e.target.value } })}>
                {shows.map((c) => (
                  <option key={c.id} value={c.id}>
                    {c.name}
                  </option>
                ))}
              </select>
            </label>
          </FieldSet>
          <FieldSet legend="Slaytları ilerletme">
            <Radio label="El ile" checked={v.advance === "manual"} onChange={() => set({ advance: "manual" })} />
            <Radio label="Varsa zamanlamaları kullanarak" checked={v.advance === "timings"} onChange={() => set({ advance: "timings" })} />
          </FieldSet>
          <FieldSet legend="Birden çok monitör">
            <Check label="Sunucu Görünümünü Kullan" checked={v.presenter} disabled={v.type !== "speaker"} onChange={(c) => set({ presenter: c })} />
            <div className="dlg-note">İzleyici gösterisi ikinci bir pencerede açılır; o pencereyi ikinci ekrana taşıyıp F11 ile tam ekran yapabilirsiniz.</div>
          </FieldSet>
        </div>
      </div>
    </Dialog>
  );
}

/** Özel Gösteriler listesi ve "Özel Gösteri Tanımla" düzenleyicisi. */
export function CustomShowsDialog({ open, onClose, actions: a, onShow }: { open: boolean; onClose: () => void; actions: Actions; onShow: (id: string) => void }) {
  const pres = a.s.pres;
  const shows = pres.customShows ?? [];
  const [selId, setSelId] = useState<string | null>(null);
  const [edit, setEdit] = useState<CustomShow | null>(null);
  const [left, setLeft] = useState<string[]>([]);
  const [right, setRight] = useState<number | null>(null);
  const sel = shows.find((c) => c.id === selId) ?? shows[0];
  const commit = (list: CustomShow[]) => a.store.commit((p) => ({ ...p, customShows: list.length ? list : undefined }));
  const saveEdit = () => {
    if (!edit) return;
    const name = edit.name.trim() || uniqueShowName(shows, "Özel Gösteri");
    const exists = shows.some((c) => c.id === edit.id);
    commit(exists ? shows.map((c) => (c.id === edit.id ? { ...edit, name } : c)) : [...shows, { ...edit, name }]);
    setSelId(edit.id);
    setEdit(null);
  };
  if (edit) {
    const moveSel = (d: -1 | 1) => {
      if (right === null) return;
      const j = right + d;
      if (j < 0 || j >= edit.slideIds.length) return;
      const ids = edit.slideIds.slice();
      [ids[right], ids[j]] = [ids[j], ids[right]];
      setEdit({ ...edit, slideIds: ids });
      setRight(j);
    };
    return (
      <Dialog open={open} title="Özel Gösteri Tanımla" onClose={() => setEdit(null)} width={640} footer={<><DialogButton primary onClick={saveEdit}>Tamam</DialogButton><DialogButton onClick={() => setEdit(null)}>İptal</DialogButton></>}>
        <label className="dlg-row">
          Slayt gösterisi adı: <input aria-label="Slayt gösterisi adı" value={edit.name} onChange={(e) => setEdit({ ...edit, name: e.target.value })} style={{ flex: 1 }} />
        </label>
        <div className="cs-cols">
          <div>
            <div className="cs-head">Sunudaki slaytlar:</div>
            <div className="cs-list oh-scroll" role="listbox" aria-label="Sunudaki slaytlar" aria-multiselectable>
              {pres.slides.map((s, i) => (
                <label key={s.id} className="cs-item">
                  <input type="checkbox" checked={left.includes(s.id)} onChange={(e) => setLeft((l) => (e.target.checked ? [...l, s.id] : l.filter((x) => x !== s.id)))} /> {i + 1}. {slideTitle(s) || "Slayt"}
                </label>
              ))}
            </div>
          </div>
          <div className="cs-mid">
            <DialogButton
              disabled={!left.length}
              onClick={() => {
                const order = pres.slides.map((s) => s.id).filter((id) => left.includes(id));
                setEdit({ ...edit, slideIds: [...edit.slideIds, ...order] });
                setLeft([]);
              }}
            >
              Ekle &gt;&gt;
            </DialogButton>
          </div>
          <div>
            <div className="cs-head">Özel gösterideki slaytlar:</div>
            <div className="cs-list oh-scroll" role="listbox" aria-label="Özel gösterideki slaytlar">
              {edit.slideIds.map((id, k) => {
                const i = pres.slides.findIndex((s) => s.id === id);
                return (
                  <div key={`${id}-${k}`} role="option" aria-selected={right === k} className={`cs-item${right === k ? " is-sel" : ""}`} onClick={() => setRight(k)}>
                    {k + 1}. {i >= 0 ? slideTitle(pres.slides[i]) || `Slayt ${i + 1}` : "(silinmiş)"}
                  </div>
                );
              })}
            </div>
          </div>
          <div className="cs-mid">
            <DialogButton disabled={right === null || right === 0} onClick={() => moveSel(-1)}>
              Yukarı
            </DialogButton>
            <DialogButton disabled={right === null || right >= edit.slideIds.length - 1} onClick={() => moveSel(1)}>
              Aşağı
            </DialogButton>
            <DialogButton
              disabled={right === null}
              onClick={() => {
                if (right === null) return;
                setEdit({ ...edit, slideIds: edit.slideIds.filter((_, k) => k !== right) });
                setRight(null);
              }}
            >
              Kaldır
            </DialogButton>
          </div>
        </div>
      </Dialog>
    );
  }
  const startEdit = (c: CustomShow) => {
    setLeft([]);
    setRight(null);
    setEdit(c);
  };
  return (
    <Dialog open={open} title="Özel Gösteriler" onClose={onClose} width={460} footer={<DialogButton primary onClick={onClose}>Kapat</DialogButton>}>
      <div className="cs-cols is-main">
        <div className="cs-list oh-scroll" role="listbox" aria-label="Özel gösteriler">
          {shows.length ? (
            shows.map((c) => (
              <div key={c.id} role="option" aria-selected={sel?.id === c.id} className={`cs-item${sel?.id === c.id ? " is-sel" : ""}`} onClick={() => setSelId(c.id)} onDoubleClick={() => startEdit(c)}>
                {c.name} ({c.slideIds.length} slayt)
              </div>
            ))
          ) : (
            <div className="cs-empty">Henüz özel gösteri yok.</div>
          )}
        </div>
        <div className="cs-mid">
          <DialogButton onClick={() => startEdit({ id: uid("cs"), name: uniqueShowName(shows, `Özel Gösteri ${shows.length + 1}`), slideIds: [] })}>Yeni...</DialogButton>
          <DialogButton disabled={!sel} onClick={() => sel && startEdit(sel)}>
            Düzenle...
          </DialogButton>
          <DialogButton disabled={!sel} onClick={() => sel && commit(shows.filter((c) => c.id !== sel.id))}>
            Kaldır
          </DialogButton>
          <DialogButton disabled={!sel} onClick={() => sel && commit([...shows, { ...sel, id: uid("cs"), name: uniqueShowName(shows, `${sel.name} kopyası`) }])}>
            Kopyala
          </DialogButton>
          <DialogButton
            disabled={!sel || !sel.slideIds.length}
            onClick={() => {
              if (!sel) return;
              onClose();
              onShow(sel.id);
            }}
          >
            Göster
          </DialogButton>
        </div>
      </div>
    </Dialog>
  );
}

/** Zamanlama Provası bitince: süreler kaydedilsin mi? */
export function RehearsalDialog({ result, onClose, actions: a }: { result: RehearsalResult | null; onClose: () => void; actions: Actions }) {
  const keep = () => {
    if (!result) return;
    a.store.commit((p) => ({
      ...p,
      slides: p.slides.map((s) => (result.times[s.id] === undefined ? s : { ...s, transition: { ...(s.transition ?? { type: "none", duration: 1, advanceOnClick: true }), advanceAfter: result.times[s.id] } })),
    }));
    a.store.set({ view: "sorter" });
    onClose();
  };
  return (
    <Dialog open={!!result} title="OfficeHere Sunum" onClose={onClose} width={420} footer={<><DialogButton primary onClick={keep}>Evet</DialogButton><DialogButton onClick={onClose}>Hayır</DialogButton></>}>
      <p style={{ margin: "4px 0 8px" }}>Slayt gösterisi {formatTime((result?.total ?? 0) * 1000, true)} sürdü. Ölçülen süreler slayt zamanlamaları olarak kaydedilsin mi?</p>
    </Dialog>
  );
}

/** Slayt boyutu değişirken içeriğin nasıl ölçekleneceği. */
export function ScaleChoiceDialog({ req, onClose, actions: a }: { req: { w: number; h: number } | null; onClose: () => void; actions: Actions }) {
  const apply = (mode: "scale" | "fit") => {
    if (req) a.store.commit((p) => resizePresentation(p, req.w, req.h, mode));
    onClose();
  };
  return (
    <Dialog open={!!req} title="Slayt Boyutu" onClose={onClose} width={460} footer={<DialogButton onClick={onClose}>İptal</DialogButton>}>
      <p style={{ margin: "4px 0 12px" }}>Slaytlar yeni boyuta göre ölçeklenecek. İçerik nasıl uyarlansın?</p>
      <div className="scale-choice">
        <button type="button" className="scale-choice-btn" onClick={() => apply("scale")}>
          <span className="scale-choice-pic is-max" aria-hidden="true" />
          En Büyük Boyuta Getir
        </button>
        <button type="button" className="scale-choice-btn" onClick={() => apply("fit")}>
          <span className="scale-choice-pic is-fit" aria-hidden="true" />
          Sığdırmayı Sağla
        </button>
      </div>
    </Dialog>
  );
}
