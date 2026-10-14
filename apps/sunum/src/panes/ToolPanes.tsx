/** Yazım Denetimi ve Slaytları Yeniden Kullan görev bölmeleri. */
import { useState } from "react";
import { openFile } from "@officehere/core";
import { OPEN_FILTERS, readPresentationFile } from "../io/formats";
import { flatShapes } from "../model/ops";
import { reuseSlide } from "../model/reuse";
import { isTextEmpty } from "../model/textStyle";
import type { Presentation } from "../model/types";
import { ScaledSlide } from "../render/SlideView";
import { slideTitle } from "../show/slideTitle";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";
import { PaneShell } from "./paneControls";

export const PROOF_LANGS: { id: string; label: string }[] = [
  { id: "tr-TR", label: "Türkçe (Türkiye)" },
  { id: "en-US", label: "İngilizce (ABD)" },
  { id: "en-GB", label: "İngilizce (Birleşik Krallık)" },
  { id: "de-DE", label: "Almanca (Almanya)" },
  { id: "fr-FR", label: "Fransızca (Fransa)" },
  { id: "es-ES", label: "İspanyolca (İspanya)" },
  { id: "it-IT", label: "İtalyanca (İtalya)" },
  { id: "ar-SA", label: "Arapça (Suudi Arabistan)" },
  { id: "ru-RU", label: "Rusça (Rusya)" },
];

export const langLabel = (id: string) => PROOF_LANGS.find((l) => l.id === id)?.label ?? id;

/** Metin içeren şekillerin (slayt, şekil) listesi — sırayla düzenlemeye açmak için. */
function textTargets(pres: Presentation): { slideIdx: number; shapeId: string }[] {
  const out: { slideIdx: number; shapeId: string }[] = [];
  pres.slides.forEach((sl, i) => {
    for (const sh of flatShapes(sl.shapes)) if (sh.kind === "sp" && sh.text && !isTextEmpty(sh.text)) out.push({ slideIdx: i, shapeId: sh.id });
  });
  return out;
}

/**
 * Yazım Denetimi: denetim tarayıcının (işletim sisteminin) yazım denetleyicisiyle
 * düzenleyicide yapılır; bu bölme dili ve açık/kapalı durumunu yönetir ve metin
 * kutularını sırayla düzenlemeye açarak hatalı sözcüklerin görülmesini sağlar.
 */
export function SpellingPane({ state, actions: a, onClose }: { state: AppState; actions: Actions; onClose: () => void }) {
  const targets = textTargets(state.pres);
  const curIdx = state.editing ? targets.findIndex((t) => t.shapeId === state.editing!.shapeId && t.slideIdx === state.slideIdx) : -1;
  const go = (dir: 1 | -1) => {
    if (!targets.length) return;
    let k = curIdx < 0 ? (dir > 0 ? targets.findIndex((t) => t.slideIdx >= state.slideIdx) : targets.length - 1) : curIdx + dir;
    if (k < 0 || k >= targets.length) k = dir > 0 ? 0 : targets.length - 1;
    const t = targets[k];
    a.goToSlide(t.slideIdx);
    a.store.set({ sel: [t.shapeId] });
    a.startEditing({ shapeId: t.shapeId });
  };
  return (
    <PaneShell title="Yazım Denetimi" onClose={onClose}>
      <div className="spell-pane">
        <label className="dlg-row">
          <input type="checkbox" checked={state.spellcheck} onChange={(e) => a.store.set({ spellcheck: e.target.checked })} /> Yazarken yazımı denetle
        </label>
        <label className="dlg-row">
          <span>Dil:</span>
          <select aria-label="Yazım denetimi dili" value={state.proofLang} onChange={(e) => a.store.set({ proofLang: e.target.value })}>
            {PROOF_LANGS.map((l) => (
              <option key={l.id} value={l.id}>
                {l.label}
              </option>
            ))}
          </select>
        </label>
        <p className="spell-note">
          Yazım hataları, metin düzenlenirken tarayıcının yazım denetleyicisiyle <span className="spell-wavy">kırmızı dalgalı</span> çizgiyle gösterilir. Öneriler için altı çizili sözcüğe <strong>Shift</strong> tuşuyla birlikte sağ tıklayın.
        </p>
        <p className="spell-note">Uygulama kendi sözlüğünü içermez; bu nedenle hatalar burada listelenmez. Aşağıdaki düğmelerle metin kutularını sırayla açarak denetleyebilirsiniz.</p>
        <div className="spell-nav">
          <button type="button" className="oh-btn" disabled={!targets.length} onClick={() => go(-1)}>
            Önceki Metin
          </button>
          <button type="button" className="oh-btn is-primary" disabled={!targets.length} onClick={() => go(1)}>
            Sonraki Metin
          </button>
        </div>
        <div className="spell-count" role="status">
          {targets.length ? `${curIdx >= 0 ? `${curIdx + 1} / ` : ""}${targets.length} metin kutusu` : "Sunuda denetlenecek metin yok."}
        </div>
      </div>
    </PaneShell>
  );
}

/** Slaytları Yeniden Kullan: başka bir sunudaki (.pptx, .odp, .ppt…) slaytları ekler. */
export function ReusePane({ actions: a, onClose }: { state: AppState; actions: Actions; onClose: () => void }) {
  const [src, setSrc] = useState<{ name: string; pres: Presentation } | null>(null);
  const [keep, setKeep] = useState(false);
  const [err, setErr] = useState("");
  const browse = async () => {
    setErr("");
    const f = await openFile(OPEN_FILTERS);
    if (!f) return;
    try {
      setSrc({ name: f.name, pres: (await readPresentationFile(f.data, f.name)).pres });
    } catch (e) {
      setErr(`Dosya açılamadı: ${(e as Error).message}`);
    }
  };
  const insert = (indices: number[]) => {
    if (!src) return;
    const dst = a.s.pres;
    const copies = indices.map((i) => reuseSlide(dst, src.pres, src.pres.slides[i], keep));
    const at = Math.min(a.s.slideIdx + 1, dst.slides.length);
    a.store.commit((p) => ({ ...p, slides: [...p.slides.slice(0, at), ...copies, ...p.slides.slice(at)] }), { slideIdx: at + copies.length - 1, sel: [], editing: null });
  };
  return (
    <PaneShell title="Slaytları Yeniden Kullan" onClose={onClose}>
      <div className="reuse-pane">
        <button type="button" className="oh-btn is-primary" onClick={() => void browse()}>
          Gözat...
        </button>
        {err && <div className="reuse-err">{err}</div>}
        {src ? (
          <>
            <div className="reuse-src" title={src.name}>
              {src.name} — {src.pres.slides.length} slayt
            </div>
            <label className="dlg-row">
              <input type="checkbox" checked={keep} onChange={(e) => setKeep(e.target.checked)} /> Kaynak biçimlendirmesini koru
            </label>
            <div className="reuse-list">
              {src.pres.slides.map((s, i) => (
                <button key={s.id} type="button" className="reuse-item" title="Slaytı ekle" onClick={() => insert([i])}>
                  <ScaledSlide pres={src.pres} slide={s} slideNumber={i + 1} width={200} mode="show" />
                  <span>
                    {i + 1}. {slideTitle(s) || "Slayt"}
                  </span>
                </button>
              ))}
            </div>
            <button type="button" className="oh-btn" onClick={() => insert(src.pres.slides.map((_, i) => i))}>
              Tüm Slaytları Ekle
            </button>
          </>
        ) : (
          <p className="spell-note">Slaytlarını eklemek istediğiniz sunuyu seçin. Eklenen slaytlar geçerli slayttan sonra yerleştirilir ve {keep ? "kaynak" : "bu sununun"} temasını kullanır.</p>
        )}
      </div>
    </PaneShell>
  );
}
