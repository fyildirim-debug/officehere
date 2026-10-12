/** Sağ bölmeler: Arka Planı Biçimlendir ve Şekli Biçimlendir. */
import { useState } from "react";
import type { Color, DashType, Fill, Line, Shape, SpShape, TextBody } from "../model/types";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";
import { readImageFile } from "../app/clipboard";
import { slideBackground } from "../render/SlideView";
import { setEffects, setFill, setLine } from "../ribbon/shapeMenus";
import { DASHES } from "../ribbon/presets";
import { cm, fromCm } from "../ribbon/common";
import { setBodyProps } from "../state/textActions";
import { NumInput, PaneColorButton, PaneShell, Radio, Row, Section, Slider } from "./paneControls";

function pickImage(cb: (src: string) => void) {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = "image/*";
  input.onchange = async () => {
    const f = input.files?.[0];
    if (f) cb((await readImageFile(f)).src);
  };
  input.click();
}

function withAlpha(c: Color, transparency: number): Color {
  const a = 1 - transparency / 100;
  return { ...c, alpha: a >= 1 ? undefined : a };
}

function FillEditor({ fill, onChange, state, allowNone = true, imageLabel = "Resim veya doku dolgusu" }: { fill: Fill; onChange: (f: Fill) => void; state: AppState; allowNone?: boolean; imageLabel?: string }) {
  const theme = state.pres.theme;
  const solidColor: Color = fill.type === "solid" ? fill.color : fill.type === "gradient" ? fill.stops[fill.stops.length - 1].color : { scheme: "accent1" };
  return (
    <>
      {allowNone && <Radio label="Dolgu yok" checked={fill.type === "none"} onChange={() => onChange({ type: "none" })} />}
      <Radio label="Düz dolgu" checked={fill.type === "solid"} onChange={() => onChange({ type: "solid", color: { ...solidColor, alpha: undefined } })} />
      <Radio
        label="Gradyan dolgu"
        checked={fill.type === "gradient"}
        onChange={() => onChange({ type: "gradient", kind: "linear", angle: 90, stops: [{ pos: 0, color: { ...solidColor, lumMod: 0.4, lumOff: 0.6 } }, { pos: 1, color: solidColor }] })}
      />
      <Radio label={imageLabel} checked={fill.type === "image"} onChange={() => pickImage((src) => onChange({ type: "image", src, mode: "stretch" }))} />
      {fill.type === "solid" && (
        <>
          <Row label="Renk">
            <PaneColorButton theme={theme} value={fill.color} onChange={(c) => c && onChange({ type: "solid", color: { ...c, alpha: fill.color.alpha } })} />
          </Row>
          <Row label="Saydamlık">
            <Slider value={Math.round((1 - (fill.color.alpha ?? 1)) * 100)} onChange={(v) => onChange({ type: "solid", color: withAlpha(fill.color, v) })} />
          </Row>
        </>
      )}
      {fill.type === "gradient" && (
        <>
          <Row label="Tür">
            <select value={fill.kind} onChange={(e) => onChange({ ...fill, kind: e.target.value as "linear" | "radial" })}>
              <option value="linear">Doğrusal</option>
              <option value="radial">Radyal</option>
            </select>
          </Row>
          {fill.kind === "linear" && (
            <Row label="Açı">
              <NumInput value={fill.angle} suffix="°" step={15} min={0} max={359} onChange={(v) => onChange({ ...fill, angle: v })} />
            </Row>
          )}
          {fill.stops.map((s, i) => (
            <Row key={i} label={`Durak ${i + 1} rengi`}>
              <PaneColorButton theme={theme} value={s.color} onChange={(c) => c && onChange({ ...fill, stops: fill.stops.map((x, k) => (k === i ? { ...x, color: c } : x)) })} />
            </Row>
          ))}
        </>
      )}
      {fill.type === "image" && (
        <>
          <Row label="Resim kaynağı">
            <button type="button" className="oh-btn" onClick={() => pickImage((src) => onChange({ ...fill, src }))}>
              Ekle...
            </button>
          </Row>
          <Row label="Döşeme">
            <input type="checkbox" checked={fill.mode === "tile"} onChange={(e) => onChange({ ...fill, mode: e.target.checked ? "tile" : "stretch" })} />
          </Row>
          <Row label="Saydamlık">
            <Slider value={Math.round((1 - (fill.alpha ?? 1)) * 100)} onChange={(v) => onChange({ ...fill, alpha: v ? 1 - v / 100 : undefined })} />
          </Row>
        </>
      )}
    </>
  );
}

export function BackgroundPane({ state, actions: a, onClose }: { state: AppState; actions: Actions; onClose: () => void }) {
  const slide = state.pres.slides[state.slideIdx];
  if (!slide) return <PaneShell title="Arka Planı Biçimlendir" onClose={onClose}>Slayt yok.</PaneShell>;
  const fill = slideBackground(state.pres, slide);
  return (
    <PaneShell title="Arka Planı Biçimlendir" onClose={onClose}>
      <Section title="Dolgu">
        <FillEditor fill={fill} state={state} allowNone={false} onChange={(f) => a.setBackground(f)} />
        <label className="fp-radio">
          <input type="checkbox" checked={!!slide.hideMasterShapes} onChange={(e) => a.setHideBackgroundGraphics(e.target.checked)} />
          Arka plan grafiklerini gizle
        </label>
      </Section>
      <div className="fp-buttons">
        <button type="button" className="oh-btn" onClick={() => a.setBackground(slide.background ?? fill, true)}>
          Tümüne Uygula
        </button>
        <button type="button" className="oh-btn" onClick={() => a.setBackground(undefined)}>
          Arka Planı Sıfırla
        </button>
      </div>
    </PaneShell>
  );
}

function firstSp(shapes: Shape[]): SpShape | undefined {
  for (const s of shapes) {
    if (s.kind === "sp") return s;
    if (s.kind === "group") {
      const c = firstSp(s.children);
      if (c) return c;
    }
  }
  return undefined;
}

export function ShapePane({ state, actions: a, onClose }: { state: AppState; actions: Actions; onClose: () => void }) {
  const [tab, setTab] = useState<"shape" | "text">("shape");
  const shapes = a.selectedShapes();
  const s = shapes[0];
  const sp = firstSp(shapes);
  const theme = state.pres.theme;
  const line: Line | undefined = s && (s.kind === "sp" || s.kind === "pic") ? s.line : undefined;
  const fx = s && (s.kind === "sp" || s.kind === "pic") ? s.effects : undefined;
  const body: TextBody | undefined = sp?.text;
  const title = s?.kind === "pic" ? "Resmi Biçimlendir" : "Şekli Biçimlendir";
  const tabs = (
    <div className="side-tabs">
      <button type="button" className={tab === "shape" ? "is-active" : ""} onClick={() => setTab("shape")}>
        Şekil Seçenekleri
      </button>
      <button type="button" className={tab === "text" ? "is-active" : ""} onClick={() => setTab("text")}>
        Metin Seçenekleri
      </button>
    </div>
  );
  if (!s)
    return (
      <PaneShell title={title} onClose={onClose}>
        <p style={{ color: "var(--oh-text-muted)" }}>Biçimlendirmek için bir nesne seçin.</p>
      </PaneShell>
    );
  const ensureLine = (l: Line): Line => (l.fill.type === "none" ? { ...l, fill: { type: "solid", color: { scheme: "accent1" } } } : l);
  return (
    <PaneShell title={title} onClose={onClose} tabs={tabs}>
      {tab === "shape" ? (
        <>
          {sp && (
            <Section title="Dolgu">
              <FillEditor fill={sp.fill} state={state} imageLabel="Resim dolgusu" onChange={(f) => setFill(a, f)} />
            </Section>
          )}
          {line && (
            <Section title={s.kind === "pic" ? "Kenarlık" : "Çizgi"}>
              <Radio label="Çizgi yok" checked={line.fill.type === "none"} onChange={() => setLine(a, (l) => ({ ...l, fill: { type: "none" } }))} />
              <Radio label="Düz çizgi" checked={line.fill.type === "solid"} onChange={() => setLine(a, ensureLine)} />
              {line.fill.type === "solid" && (
                <>
                  <Row label="Renk">
                    <PaneColorButton theme={theme} value={line.fill.color} onChange={(c) => c && setLine(a, (l) => ({ ...l, fill: { type: "solid", color: c } }))} />
                  </Row>
                  <Row label="Saydamlık">
                    <Slider value={Math.round((1 - (line.fill.color.alpha ?? 1)) * 100)} onChange={(v) => setLine(a, (l) => (l.fill.type === "solid" ? { ...l, fill: { type: "solid", color: withAlpha(l.fill.color, v) } } : l))} />
                  </Row>
                  <Row label="Genişlik">
                    <NumInput value={line.width} suffix="nk" step={0.25} min={0} max={1584} onChange={(v) => setLine(a, (l) => ({ ...l, width: v }))} />
                  </Row>
                  <Row label="Tire türü">
                    <select value={line.dash} onChange={(e) => setLine(a, (l) => ({ ...l, dash: e.target.value as DashType }))}>
                      {DASHES.map((d) => (
                        <option key={d.id} value={d.id}>
                          {d.label}
                        </option>
                      ))}
                    </select>
                  </Row>
                  {sp && (
                    <>
                      <Row label="Başlangıç ok türü">
                        <select value={line.head ?? "none"} onChange={(e) => setLine(a, (l) => ({ ...l, head: e.target.value === "none" ? undefined : (e.target.value as Line["head"]) }))}>
                          {["none", "triangle", "arrow", "stealth", "diamond", "oval"].map((t) => (
                            <option key={t} value={t}>
                              {ARROW_NAMES[t]}
                            </option>
                          ))}
                        </select>
                      </Row>
                      <Row label="Bitiş ok türü">
                        <select value={line.tail ?? "none"} onChange={(e) => setLine(a, (l) => ({ ...l, tail: e.target.value === "none" ? undefined : (e.target.value as Line["tail"]) }))}>
                          {["none", "triangle", "arrow", "stealth", "diamond", "oval"].map((t) => (
                            <option key={t} value={t}>
                              {ARROW_NAMES[t]}
                            </option>
                          ))}
                        </select>
                      </Row>
                    </>
                  )}
                </>
              )}
            </Section>
          )}
          <Section title="Efektler" open={false}>
            <label className="fp-radio">
              <input type="checkbox" checked={!!fx?.shadow} onChange={(e) => setEffects(a, (f) => ({ ...f, shadow: e.target.checked ? { color: { rgb: "#000000", alpha: 0.4 }, blur: 5, dist: 3, dir: 45 } : undefined }))} />
              Gölge
            </label>
            {fx?.shadow && (
              <>
                <Row label="Renk">
                  <PaneColorButton theme={theme} value={fx.shadow.color} onChange={(c) => c && setEffects(a, (f) => ({ ...f, shadow: f.shadow && { ...f.shadow, color: { ...c, alpha: f.shadow.color.alpha } } }))} />
                </Row>
                <Row label="Saydamlık">
                  <Slider value={Math.round((1 - (fx.shadow.color.alpha ?? 1)) * 100)} onChange={(v) => setEffects(a, (f) => ({ ...f, shadow: f.shadow && { ...f.shadow, color: withAlpha(f.shadow.color, v) } }))} />
                </Row>
                <Row label="Bulanıklaştır">
                  <NumInput value={fx.shadow.blur * 0.75} suffix="nk" min={0} max={100} onChange={(v) => setEffects(a, (f) => ({ ...f, shadow: f.shadow && { ...f.shadow, blur: v / 0.75 } }))} />
                </Row>
                <Row label="Uzaklık">
                  <NumInput value={fx.shadow.dist * 0.75} suffix="nk" min={0} max={200} onChange={(v) => setEffects(a, (f) => ({ ...f, shadow: f.shadow && { ...f.shadow, dist: v / 0.75 } }))} />
                </Row>
                <Row label="Açı">
                  <NumInput value={fx.shadow.dir} suffix="°" min={0} max={359} step={15} onChange={(v) => setEffects(a, (f) => ({ ...f, shadow: f.shadow && { ...f.shadow, dir: v } }))} />
                </Row>
              </>
            )}
            <label className="fp-radio">
              <input type="checkbox" checked={!!fx?.glow} onChange={(e) => setEffects(a, (f) => ({ ...f, glow: e.target.checked ? { color: { scheme: "accent1", alpha: 0.6 }, radius: 10 } : undefined }))} />
              Parlama
            </label>
            {fx?.glow && (
              <>
                <Row label="Renk">
                  <PaneColorButton theme={theme} value={fx.glow.color} onChange={(c) => c && setEffects(a, (f) => ({ ...f, glow: f.glow && { ...f.glow, color: { ...c, alpha: 0.6 } } }))} />
                </Row>
                <Row label="Boyut">
                  <NumInput value={fx.glow.radius * 0.75} suffix="nk" min={0} max={150} onChange={(v) => setEffects(a, (f) => ({ ...f, glow: f.glow && { ...f.glow, radius: v / 0.75 } }))} />
                </Row>
              </>
            )}
            <label className="fp-radio">
              <input type="checkbox" checked={!!fx?.reflection} onChange={(e) => setEffects(a, (f) => ({ ...f, reflection: e.target.checked || undefined }))} />
              Yansıma
            </label>
          </Section>
          <Section title="Boyut ve Özellikler">
            <Row label="Yükseklik">
              <NumInput value={cm(s.h)} suffix="cm" step={0.1} min={0} onChange={(v) => a.setSize(undefined, fromCm(v), s.kind === "pic")} />
            </Row>
            <Row label="Genişlik">
              <NumInput value={cm(s.w)} suffix="cm" step={0.1} min={0} onChange={(v) => a.setSize(fromCm(v), undefined, s.kind === "pic")} />
            </Row>
            <Row label="Döndürme">
              <NumInput value={s.rot} suffix="°" step={1} onChange={(v) => a.setRotation(v)} />
            </Row>
            <Row label="Yatay konum">
              <NumInput value={cm(s.x)} suffix="cm" step={0.1} onChange={(v) => a.updateShapes((x) => ({ ...x, ...shift(x, fromCm(v) - x.x, 0) }) as Shape)} />
            </Row>
            <Row label="Dikey konum">
              <NumInput value={cm(s.y)} suffix="cm" step={0.1} onChange={(v) => a.updateShapes((x) => ({ ...x, ...shift(x, 0, fromCm(v) - x.y) }) as Shape)} />
            </Row>
          </Section>
        </>
      ) : sp ? (
        <Section title="Metin Kutusu">
          <Row label="Dikey hizalama">
            <select value={body?.anchor ?? "t"} onChange={(e) => setBodyProps(a, { anchor: e.target.value as TextBody["anchor"] })}>
              <option value="t">Üst</option>
              <option value="ctr">Orta</option>
              <option value="b">Alt</option>
            </select>
          </Row>
          <Row label="Metin yönü">
            <select value={body?.vert ?? "horz"} onChange={(e) => setBodyProps(a, { vert: e.target.value as TextBody["vert"] })}>
              <option value="horz">Yatay</option>
              <option value="vert">Tüm metni 90° döndür</option>
              <option value="vert270">Tüm metni 270° döndür</option>
              <option value="stacked">Yığılmış</option>
            </select>
          </Row>
          <Radio label="Otomatik Sığdırma" checked={(body?.autofit ?? "none") === "none"} onChange={() => setBodyProps(a, { autofit: "none", fontScale: undefined, lnSpcReduction: undefined })} />
          <Radio label="Taşarsa metni küçült" checked={body?.autofit === "shrink"} onChange={() => setBodyProps(a, { autofit: "shrink" })} />
          <Radio label="Şekli metne göre yeniden boyutlandır" checked={body?.autofit === "shape"} onChange={() => setBodyProps(a, { autofit: "shape", fontScale: undefined })} />
          {(["l", "r", "t", "b"] as const).map((k) => (
            <Row key={k} label={{ l: "Sol kenar boşluğu", r: "Sağ kenar boşluğu", t: "Üst kenar boşluğu", b: "Alt kenar boşluğu" }[k]}>
              <NumInput
                value={cm((body?.insets ?? { l: 9.6, t: 4.8, r: 9.6, b: 4.8 })[k])}
                suffix="cm"
                step={0.05}
                min={0}
                onChange={(v) => setBodyProps(a, { insets: { ...(body?.insets ?? { l: 9.6, t: 4.8, r: 9.6, b: 4.8 }), [k]: fromCm(v) } })}
              />
            </Row>
          ))}
          <label className="fp-radio">
            <input type="checkbox" checked={body?.wrap ?? true} onChange={(e) => setBodyProps(a, { wrap: e.target.checked })} />
            Metni şekilde kaydır
          </label>
        </Section>
      ) : (
        <p style={{ color: "var(--oh-text-muted)" }}>Seçili nesnede metin yok.</p>
      )}
    </PaneShell>
  );
}

const ARROW_NAMES: Record<string, string> = { none: "Yok", triangle: "Ok", arrow: "Açık Ok", stealth: "Gizli Ok", diamond: "Elmas Ok", oval: "Oval Ok" };

function shift(s: Shape, dx: number, dy: number): Partial<Shape> {
  if (s.kind === "group") return { x: s.x + dx, y: s.y + dy, children: s.children.map((c) => ({ ...c, ...shift(c, dx, dy) }) as Shape) } as Partial<Shape>;
  return { x: s.x + dx, y: s.y + dy };
}
