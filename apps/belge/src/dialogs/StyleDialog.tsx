// Biçimlendirmeden Yeni Stil Oluştur / Stili Değiştir: ad, tür, temel stil, sonraki paragraf stili, yazı tipi,
// paragraf biçimi ve önizleme. Stiller belge ayarlarında (customStyles) tutulur ve DOCX'e yazılır.
import { useMemo, useState } from "react";
import { TextAlignCenter20Regular, TextAlignJustify20Regular, TextAlignLeft20Regular, TextAlignRight20Regular } from "@fluentui/react-icons";
import { FONT_LIST, FONT_SIZES, fontStack } from "@officehere/core";
import { Dialog, DialogButton, Field, FieldSet } from "@officehere/ui";
import { useBelge } from "../state/context";
import { applyStyle, currentPara } from "../editor/commands";
import { LINE_FACTOR, effectiveProps, styleIdFromName, styleIdOf, styleIds, type StyleDef, type StyleProps } from "../model/styles";
import { Check } from "./FontDialog";

const PT_PER_CM = 72 / 2.54;
const num = (s: string) => {
  const n = parseFloat(s.replace(",", "."));
  return Number.isFinite(n) ? n : 0;
};

export function StyleDialog({ styleId, onClose }: { styleId?: string; onClose: () => void }) {
  const ctx = useBelge();
  const { styles, editor, fmt } = ctx;
  const editing = styleId ? styles[styleId] : undefined;
  // Yeni stil seçimin biçimiyle başlar (Word: "Biçimlendirmeden Yeni Stil Oluştur").
  const init = useMemo<StyleProps>(() => {
    if (editing) return effectiveProps(styles, editing.id);
    const p = currentPara(editor);
    const base = effectiveProps(styles, p ? styleIdOf(p.node.type.name, p.node.attrs) : "Normal");
    return { ...base, font: fmt.font || base.font, size: fmt.size || base.size, bold: fmt.bold || undefined, italic: fmt.italic || undefined, underline: !!fmt.underline || undefined };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const [name, setName] = useState(editing?.name ?? `Stil${Object.keys(ctx.settings.customStyles ?? {}).length + 1}`);
  const [type, setType] = useState<StyleDef["type"]>(editing?.type ?? "paragraph");
  const [basedOn, setBasedOn] = useState(editing?.basedOn ?? (editing?.id === "Normal" ? "" : "Normal"));
  const [next, setNext] = useState(editing?.next ?? editing?.id ?? "");
  const [p, setP] = useState<StyleProps>(init);
  const [gallery, setGallery] = useState(editing ? editing.gallery !== false : true);
  const set = (patch: Partial<StyleProps>) => setP((x) => ({ ...x, ...patch }));
  const paraStyles = styleIds(styles).filter((id) => styles[id].type === "paragraph" && id !== editing?.id);
  const nameTaken = !editing && Object.values(styles).some((s) => s.name.toLocaleLowerCase("tr-TR") === name.trim().toLocaleLowerCase("tr-TR"));

  const ok = () => {
    if (!name.trim() || nameTaken) return;
    const id = editing?.id ?? styleIdFromName(name.trim(), styles);
    // Temel stilden farklı olan özellikler saklanır.
    const base = basedOn ? effectiveProps(styles, basedOn) : {};
    const props: StyleProps = {};
    for (const [k, v] of Object.entries(p) as [keyof StyleProps, unknown][]) {
      if (v === undefined) continue;
      if (JSON.stringify(v) !== JSON.stringify((base as Record<string, unknown>)[k]) || id === "Normal") (props as Record<string, unknown>)[k] = v;
    }
    const def: StyleDef = { id, name: name.trim(), type, props, ...(basedOn && id !== "Normal" ? { basedOn } : {}), ...(type === "paragraph" && next ? { next } : {}), ...(editing?.level ? { level: editing.level } : {}), gallery };
    ctx.updateSettings((s) => ({ ...s, customStyles: { ...(s.customStyles ?? {}), [id]: def } }));
    if (!editing) requestAnimationFrame(() => applyStyle(editor, { ...styles, [id]: def }, id));
    onClose();
  };

  const preview: React.CSSProperties = {
    fontFamily: p.font ? fontStack(p.font) : undefined,
    fontSize: `${p.size ?? 11}pt`,
    fontWeight: p.bold ? 700 : 400,
    fontStyle: p.italic ? "italic" : "normal",
    textDecoration: p.underline ? "underline" : undefined,
    color: p.color,
    textAlign: p.align,
    lineHeight: p.line ? p.line * LINE_FACTOR : undefined,
    paddingLeft: p.indentLeft ? `${p.indentLeft}pt` : undefined,
  };

  return (
    <Dialog
      open
      title={editing ? "Stili Değiştir" : "Biçimlendirmeden Yeni Stil Oluştur"}
      width={560}
      onClose={onClose}
      onSubmit={ok}
      footer={
        <>
          <DialogButton primary onClick={ok} disabled={!name.trim() || nameTaken}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <FieldSet legend="Özellikler">
        <div className="bl-dlg-grid">
          <Field label="Ad:">
            <input type="text" aria-label="Ad" value={name} onChange={(e) => setName(e.target.value)} disabled={!!editing && editing.id in BUILTIN_LOCKED} />
          </Field>
          <Field label="Stil türü:">
            <select aria-label="Stil türü" value={type} disabled={!!editing} onChange={(e) => setType(e.target.value as StyleDef["type"])}>
              <option value="paragraph">Paragraf</option>
              <option value="character">Karakter</option>
            </select>
          </Field>
          <Field label="Stil temeli:">
            <select aria-label="Stil temeli" value={basedOn} disabled={editing?.id === "Normal"} onChange={(e) => setBasedOn(e.target.value)}>
              <option value="">(stil yok)</option>
              {paraStyles.map((id) => (
                <option key={id} value={id}>
                  {styles[id].name}
                </option>
              ))}
            </select>
          </Field>
          <Field label="Sonraki paragraf için stil:">
            <select aria-label="Sonraki paragraf için stil" value={next} disabled={type !== "paragraph"} onChange={(e) => setNext(e.target.value)}>
              <option value="">(aynı stil)</option>
              {styleIds(styles)
                .filter((id) => styles[id].type === "paragraph")
                .map((id) => (
                  <option key={id} value={id}>
                    {styles[id].name}
                  </option>
                ))}
            </select>
          </Field>
        </div>
        {nameTaken && <p className="bl-pr-error">Bu adda bir stil zaten var.</p>}
      </FieldSet>
      <FieldSet legend="Biçimlendirme">
        <div className="bl-style-fmt">
          <select aria-label="Yazı tipi" value={p.font ?? ""} onChange={(e) => set({ font: e.target.value })}>
            {FONT_LIST.map((f) => (
              <option key={f} value={f}>
                {f}
              </option>
            ))}
          </select>
          <select aria-label="Yazı tipi boyutu" value={String(p.size ?? 11)} onChange={(e) => set({ size: num(e.target.value) })}>
            {FONT_SIZES.map((s) => (
              <option key={s} value={String(s)}>
                {String(s).replace(".", ",")}
              </option>
            ))}
          </select>
          <button type="button" className={`bl-tog${p.bold ? " is-on" : ""}`} aria-pressed={!!p.bold} aria-label="Kalın" onClick={() => set({ bold: !p.bold })}>
            <b>K</b>
          </button>
          <button type="button" className={`bl-tog${p.italic ? " is-on" : ""}`} aria-pressed={!!p.italic} aria-label="İtalik" onClick={() => set({ italic: !p.italic })}>
            <i>T</i>
          </button>
          <button type="button" className={`bl-tog${p.underline ? " is-on" : ""}`} aria-pressed={!!p.underline} aria-label="Altı Çizili" onClick={() => set({ underline: !p.underline })}>
            <u>A</u>
          </button>
          <input type="color" aria-label="Yazı tipi rengi" value={p.color ?? "#000000"} onChange={(e) => set({ color: e.target.value.toUpperCase() })} />
        </div>
        {type === "paragraph" && (
          <div className="bl-style-fmt">
            {(["left", "center", "right", "justify"] as const).map((al) => (
              <button key={al} type="button" className={`bl-tog${(p.align ?? "left") === al ? " is-on" : ""}`} aria-label={ALIGN_LABEL[al]} onClick={() => set({ align: al })}>
                {ALIGN_ICON[al]}
              </button>
            ))}
            <label>
              Satır aralığı:
              <select aria-label="Satır aralığı" value={String(p.line ?? 1.08)} onChange={(e) => set({ line: num(e.target.value) })}>
                {["1", "1.08", "1.15", "1.5", "2"].map((v) => (
                  <option key={v} value={v}>
                    {v.replace(".", ",")}
                  </option>
                ))}
              </select>
            </label>
            <label>
              Önce:
              <input type="number" aria-label="Önce" min={0} step={6} value={p.spaceBefore ?? 0} onChange={(e) => set({ spaceBefore: num(e.target.value) })} />
            </label>
            <label>
              Sonra:
              <input type="number" aria-label="Sonra" min={0} step={6} value={p.spaceAfter ?? 0} onChange={(e) => set({ spaceAfter: num(e.target.value) })} />
            </label>
            <label>
              Sol girinti (cm):
              <input type="number" aria-label="Sol girinti" min={0} step={0.25} value={+((p.indentLeft ?? 0) / PT_PER_CM).toFixed(2)} onChange={(e) => set({ indentLeft: +(num(e.target.value) * PT_PER_CM).toFixed(2) })} />
            </label>
          </div>
        )}
        <div className="bl-style-preview">
          <span className="bl-style-prev-dim">Önceki Paragraf Önceki Paragraf Önceki Paragraf</span>
          <div style={preview}>Örnek Metin Örnek Metin Örnek Metin — İıĞğŞşÜüÖöÇç</div>
          <span className="bl-style-prev-dim">Sonraki Paragraf Sonraki Paragraf Sonraki Paragraf</span>
        </div>
      </FieldSet>
      <Check label="Stil galerisine ekle" v={gallery} set={setGallery} />
    </Dialog>
  );
}

const BUILTIN_LOCKED: Record<string, true> = { Normal: true, Heading1: true, Heading2: true, Heading3: true, Heading4: true, Heading5: true, Heading6: true };
const ALIGN_LABEL = { left: "Sola Hizala", center: "Ortala", right: "Sağa Hizala", justify: "İki Yana Yasla" } as const;
const ALIGN_ICON = { left: <TextAlignLeft20Regular />, center: <TextAlignCenter20Regular />, right: <TextAlignRight20Regular />, justify: <TextAlignJustify20Regular /> } as const;
