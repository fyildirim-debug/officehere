/** Eylem Ayarları, Yakınlaştırma (Slayt / Özet) ve Resimleri Sıkıştır iletişim kutuları. */
import { useState } from "react";
import { Dialog, DialogButton, DialogTabs, Field } from "@officehere/ui";
import { JUMP_LABELS, normalizeUrl, parseLink } from "../model/links";
import type { Presentation, Shape } from "../model/types";
import { ScaledSlide } from "../render/SlideView";
import { slideTitle } from "../show/SlideShow";
import type { Actions } from "../state/actions";
import { compressPictures, COMPRESS_PPI, type CompressOptions } from "../state/compress";
import { insertSlideZoom, insertSummaryZoom } from "../state/zoomActions";

type ActionKind = "none" | "next" | "prev" | "first" | "last" | "slide" | "url";

function toKind(link: string | undefined): { kind: ActionKind; slideId: string; url: string } {
  if (!link) return { kind: "none", slideId: "", url: "" };
  const t = parseLink(link);
  if (t.kind === "jump") return { kind: t.jump, slideId: "", url: "" };
  if (t.kind === "slide") return { kind: "slide", slideId: t.slideId, url: "" };
  return { kind: "url", slideId: "", url: t.url };
}

function fromKind(k: { kind: ActionKind; slideId: string; url: string }): string | undefined {
  if (k.kind === "none") return undefined;
  if (k.kind === "slide") return k.slideId ? `#slide:${k.slideId}` : undefined;
  if (k.kind === "url") return normalizeUrl(k.url) || undefined;
  return `#${k.kind}`;
}

function ActionPanel({ pres, value, onChange }: { pres: Presentation; value: { kind: ActionKind; slideId: string; url: string }; onChange: (v: { kind: ActionKind; slideId: string; url: string }) => void }) {
  const link = value.kind !== "none";
  return (
    <div className="action-panel">
      <label className="oh-radio">
        <input type="radio" checked={!link} onChange={() => onChange({ ...value, kind: "none" })} /> Yok
      </label>
      <label className="oh-radio">
        <input type="radio" checked={link} onChange={() => onChange({ ...value, kind: "next" })} /> Köprü:
      </label>
      <select aria-label="Köprü hedefi" disabled={!link} value={link ? value.kind : "next"} onChange={(e) => onChange({ ...value, kind: e.target.value as ActionKind, slideId: value.slideId || pres.slides[0]?.id || "" })} style={{ width: "100%", height: 26, marginLeft: 22, maxWidth: 300 }}>
        {(["next", "prev", "first", "last"] as const).map((j) => (
          <option key={j} value={j}>
            {JUMP_LABELS[j]}
          </option>
        ))}
        <option value="slide">Slayt...</option>
        <option value="url">URL...</option>
      </select>
      {value.kind === "slide" && (
        <Field label="Slayt başlığı:">
          <select aria-label="Slayt" size={6} value={value.slideId} onChange={(e) => onChange({ ...value, slideId: e.target.value })} style={{ width: "100%" }}>
            {pres.slides.map((s, i) => (
              <option key={s.id} value={s.id}>
                {i + 1}. {slideTitle(s) || "Slayt"}
              </option>
            ))}
          </select>
        </Field>
      )}
      {value.kind === "url" && (
        <Field label="URL:">
          <input type="text" aria-label="URL" value={value.url} placeholder="https://" onChange={(e) => onChange({ ...value, url: e.target.value })} style={{ width: "100%", height: 26 }} />
        </Field>
      )}
    </div>
  );
}

/** Eylem Ayarları: şekle tıklanınca / üzerine gelinince köprü (sonraki/önceki/ilk/son slayt, slayt, URL). */
export function ActionDialog({ open, onClose, actions: a }: { open: boolean; onClose: () => void; actions: Actions }) {
  const shape: Shape | undefined = a.selectedShapes()[0];
  const [tab, setTab] = useState("click");
  const [click, setClick] = useState(() => toKind(shape?.link));
  const [hover, setHover] = useState(() => toKind(shape?.hoverLink));
  const [last, setLast] = useState(false);
  if (open !== last) {
    setLast(open);
    if (open) {
      setTab("click");
      setClick(toKind(shape?.link));
      setHover(toKind(shape?.hoverLink));
    }
  }
  if (!open) return null;
  const ok = () => {
    if (shape) {
      const link = fromKind(click);
      const hoverLink = fromKind(hover);
      a.updateShapes((s) => ({ ...s, link, hoverLink }), [shape.id]);
    }
    onClose();
  };
  return (
    <Dialog
      open
      title="Eylem Ayarları"
      onClose={onClose}
      onSubmit={ok}
      width={400}
      footer={
        <>
          <DialogButton primary onClick={ok}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      {!shape ? (
        <p>Eylem eklemek için önce bir şekil, resim ya da metin kutusu seçin.</p>
      ) : (
        <>
          <DialogTabs
            tabs={[
              { id: "click", label: "Fare Tıklaması" },
              { id: "hover", label: "Fare Üzerine Gelme" },
            ]}
            active={tab}
            onChange={setTab}
          />
          <div style={{ marginTop: 10 }}>
            <div style={{ marginBottom: 6, fontWeight: 600 }}>{tab === "click" ? "Tıklandığında yapılacak eylem" : "Fare üzerine gelince yapılacak eylem"}</div>
            {tab === "click" ? <ActionPanel pres={a.pres} value={click} onChange={setClick} /> : <ActionPanel pres={a.pres} value={hover} onChange={setHover} />}
          </div>
        </>
      )}
    </Dialog>
  );
}

/** Yakınlaştırma Ekle: slaytları seçme ızgarası (Slayt Yakınlaştırma / Özet Yakınlaştırma). */
export function InsertZoomDialog({ kind, onClose, actions: a }: { kind: "slide" | "summary" | null; onClose: () => void; actions: Actions }) {
  const [sel, setSel] = useState<string[]>([]);
  const [last, setLast] = useState<typeof kind>(null);
  const [busy, setBusy] = useState(false);
  if (kind !== last) {
    setLast(kind);
    setSel([]);
  }
  if (!kind) return null;
  const pres = a.pres;
  const toggle = (id: string) => setSel((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
  const ok = async () => {
    if (!sel.length || busy) return;
    setBusy(true);
    try {
      if (kind === "slide") await insertSlideZoom(a, sel);
      else await insertSummaryZoom(a, sel);
    } finally {
      setBusy(false);
    }
    onClose();
  };
  return (
    <Dialog
      open
      title={kind === "slide" ? "Slayt Yakınlaştırma Ekle" : "Özet Yakınlaştırma Ekle"}
      onClose={onClose}
      width={620}
      footer={
        <>
          <DialogButton primary onClick={() => void ok()} disabled={!sel.length || busy}>
            Ekle
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <p style={{ margin: "0 0 8px" }}>{kind === "slide" ? "Yakınlaştırmasını eklemek istediğiniz slaytları seçin." : "Özete eklenecek slaytları seçin; her biri bir bölümün ilk slaydı olur."}</p>
      <div className="zoom-pick-grid" role="listbox" aria-multiselectable="true" aria-label="Slaytlar">
        {pres.slides.map((s, i) => (
          <label key={s.id} className={`zoom-pick${sel.includes(s.id) ? " is-selected" : ""}`} role="option" aria-selected={sel.includes(s.id)}>
            <span className="thumb-box" style={{ width: 128 }}>
              <ScaledSlide pres={pres} slide={s} slideNumber={i + 1} width={128} />
            </span>
            <span className="zoom-pick-cap">
              <input type="checkbox" checked={sel.includes(s.id)} onChange={() => toggle(s.id)} /> {i + 1}. {slideTitle(s) || "Slayt"}
            </span>
          </label>
        ))}
      </div>
    </Dialog>
  );
}

/** Resimleri Sıkıştır: çözünürlüğü düşürür (tuvalle yeniden örnekleme) ve kırpılan alanları siler. */
export function CompressDialog({ open, onClose, actions: a }: { open: boolean; onClose: () => void; actions: Actions }) {
  const [opts, setOpts] = useState<CompressOptions>({ onlySelected: true, deleteCrop: true, ppi: 150 });
  const [busy, setBusy] = useState(false);
  if (!open) return null;
  const hasSel = a.selectedShapes().some((s) => s.kind === "pic");
  const ok = async () => {
    setBusy(true);
    try {
      const r = await compressPictures(a, { ...opts, onlySelected: opts.onlySelected && hasSel });
      a.message(r.count ? `${r.count} resim sıkıştırıldı (${Math.round(r.before / 1024)} KB → ${Math.round(r.after / 1024)} KB).` : "Sıkıştırılacak resim bulunamadı.", r.count ? "success" : "info");
    } finally {
      setBusy(false);
    }
    onClose();
  };
  return (
    <Dialog
      open
      title="Resimleri Sıkıştır"
      onClose={onClose}
      onSubmit={() => void ok()}
      width={380}
      footer={
        <>
          <DialogButton primary onClick={() => void ok()} disabled={busy}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <div style={{ fontWeight: 600, marginBottom: 4 }}>Sıkıştırma seçenekleri:</div>
      <label className="oh-check">
        <input type="checkbox" checked={opts.onlySelected && hasSel} disabled={!hasSel} onChange={(e) => setOpts({ ...opts, onlySelected: e.target.checked })} /> Yalnızca bu resme uygula
      </label>
      <label className="oh-check">
        <input type="checkbox" checked={opts.deleteCrop} onChange={(e) => setOpts({ ...opts, deleteCrop: e.target.checked })} /> Resimlerin kırpılan alanlarını sil
      </label>
      <div style={{ fontWeight: 600, margin: "10px 0 4px" }}>Çözünürlük:</div>
      {COMPRESS_PPI.map((c) => (
        <label key={c.ppi} className="oh-radio" style={{ display: "block" }}>
          <input type="radio" name="ppi" checked={opts.ppi === c.ppi} onChange={() => setOpts({ ...opts, ppi: c.ppi })} /> {c.label}
        </label>
      ))}
    </Dialog>
  );
}
