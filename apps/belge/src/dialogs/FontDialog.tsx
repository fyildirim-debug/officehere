// Yazı Tipi iletişim kutusu (Ctrl+D).
import { useState } from "react";
import { FONT_LIST, FONT_SIZES, fontStack } from "@officehere/core";
import { Dialog, DialogButton, DialogTabs, Field, FieldSet } from "@officehere/ui";
import { useBelge } from "../state/context";
import { AutoFocus } from "./AutoFocus";
import { UNDERLINE_TYPES, type UnderlineType } from "../editor/marks";

const STYLES = ["Normal", "İtalik", "Kalın", "Kalın İtalik"];

export function FontDialog({ onClose }: { onClose: () => void }) {
  const ctx = useBelge();
  const { editor, fmt } = ctx;
  const ts = editor.getAttributes("textStyle");
  const [tab, setTab] = useState("font");
  const [font, setFont] = useState(fmt.font);
  const [size, setSize] = useState(String(fmt.size).replace(".", ","));
  const [style, setStyle] = useState(STYLES[(fmt.italic ? 1 : 0) + (fmt.bold ? 2 : 0)]);
  const [color, setColor] = useState<string | null>(fmt.color);
  const [uType, setUType] = useState<string>(fmt.underline ?? "none");
  const [uColor, setUColor] = useState<string | null>((editor.getAttributes("underline").color as string | null) ?? null);
  const [strike, setStrike] = useState(fmt.strike);
  const [dstrike, setDstrike] = useState(!!ts.doubleStrike);
  const [sup, setSup] = useState(fmt.sup);
  const [sub, setSub] = useState(fmt.sub);
  const [smallCaps, setSmallCaps] = useState(!!ts.smallCaps);
  const [caps, setCaps] = useState(!!ts.caps);
  const [hidden, setHidden] = useState(!!ts.hidden);
  const ls = (ts.letterSpacing as number | null) ?? 0;
  const [spacingKind, setSpacingKind] = useState(ls > 0 ? "expanded" : ls < 0 ? "condensed" : "normal");
  const [spacing, setSpacing] = useState(String(Math.abs(ls)).replace(".", ","));
  const [effects, setEffects] = useState({ shadow: !!ts.shadow, outline: !!ts.outline, emboss: !!ts.emboss, imprint: !!ts.imprint });
  const [showEffects, setShowEffects] = useState(false);

  const sizeNum = parseFloat(size.replace(",", ".")) || fmt.size;
  const bold = style.includes("Kalın");
  const italic = style.includes("İtalik");
  const spacingVal = spacingKind === "normal" ? 0 : (parseFloat(spacing.replace(",", ".")) || 1) * (spacingKind === "condensed" ? -1 : 1);

  const apply = () => {
    let c = editor.chain().focus().setFontFamily(font).setFontSize(`${sizeNum}pt`);
    c = bold ? c.setBold() : c.unsetBold();
    c = italic ? c.setItalic() : c.unsetItalic();
    c = color ? c.setColor(color) : c.unsetColor();
    c = uType === "none" ? c.unsetUnderline() : c.setUnderlineType({ type: uType as UnderlineType, color: uColor });
    c = strike ? c.setStrike() : c.unsetStrike();
    c = sup ? c.setSuperscript() : c.unsetSuperscript();
    c = sub ? c.setSubscript() : c.unsetSubscript();
    c = c.setMark("textStyle", { caps, smallCaps, hidden, doubleStrike: dstrike, letterSpacing: spacingVal || null, ...effects });
    c.run();
    onClose();
  };

  /** Varsayılan Olarak Ayarla: bu belgedeki Normal stilin yazı tipi. */
  const setDefault = () => {
    const normal = ctx.styles.Normal;
    ctx.updateSettings((s) => ({
      ...s,
      customStyles: { ...(s.customStyles ?? {}), Normal: { ...normal, props: { ...normal.props, font, size: sizeNum, bold: bold || undefined, italic: italic || undefined, color: color ?? undefined } } },
    }));
    ctx.notify("Yazı tipi bu belgenin varsayılanı (Normal stili) olarak ayarlandı.", "success");
    onClose();
  };

  const preview: React.CSSProperties = {
    fontFamily: fontStack(font),
    fontSize: `${Math.min(sizeNum, 28)}pt`,
    fontWeight: bold ? 700 : 400,
    fontStyle: italic ? "italic" : "normal",
    color: color ?? "#000",
    textDecorationLine: [uType !== "none" ? "underline" : "", strike || dstrike ? "line-through" : ""].filter(Boolean).join(" ") || "none",
    textDecorationStyle: uType === "double" || dstrike ? "double" : uType === "dotted" ? "dotted" : uType === "dashed" ? "dashed" : uType === "wave" ? "wavy" : "solid",
    textDecorationColor: uColor ?? undefined,
    textTransform: caps ? "uppercase" : undefined,
    fontVariant: smallCaps ? "small-caps" : undefined,
    letterSpacing: spacingVal ? `${spacingVal}pt` : undefined,
    verticalAlign: sup ? "super" : sub ? "sub" : undefined,
    textShadow: effects.shadow ? "1px 1px 1px rgba(0,0,0,0.45)" : effects.emboss ? "-1px -1px 0 #fff, 1px 1px 0 rgba(0,0,0,0.35)" : effects.imprint ? "1px 1px 0 #fff, -1px -1px 0 rgba(0,0,0,0.35)" : undefined,
    WebkitTextStroke: effects.outline ? `0.6px ${color ?? "#000"}` : undefined,
    WebkitTextFillColor: effects.outline ? "#fff" : undefined,
  };

  return (
    <Dialog
      open
      title="Yazı Tipi"
      width={520}
      onClose={onClose}
      onSubmit={apply}
      footer={
        <>
          <DialogButton onClick={setDefault}>Varsayılan Olarak Ayarla</DialogButton>
          <DialogButton onClick={() => setShowEffects((v) => !v)}>Metin Efektleri...</DialogButton>
          <DialogButton primary onClick={apply}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      {showEffects && (
        <FieldSet legend="Metin Efektleri">
          <div className="bl-dlg-checks">
            <Check label="Gölge" v={effects.shadow} set={(v) => setEffects((e) => ({ ...e, shadow: v }))} />
            <Check label="Anahat" v={effects.outline} set={(v) => setEffects((e) => ({ ...e, outline: v }))} />
            <Check label="Kabartma" v={effects.emboss} set={(v) => setEffects((e) => ({ ...e, emboss: v, imprint: v ? false : e.imprint }))} />
            <Check label="Oyma" v={effects.imprint} set={(v) => setEffects((e) => ({ ...e, imprint: v, emboss: v ? false : e.emboss }))} />
          </div>
        </FieldSet>
      )}
      <DialogTabs
        tabs={[
          { id: "font", label: "Yazı Tipi" },
          { id: "adv", label: "Gelişmiş" },
        ]}
        active={tab}
        onChange={setTab}
      />
      <AutoFocus deps={[tab]} />
      {tab === "font" ? (
        <>
          <div className="bl-dlg-cols3">
            <div className="bl-listfield">
              <span>Yazı tipi:</span>
              <input value={font} onChange={(e) => setFont(e.target.value)} />
              <select size={6} value={font} onChange={(e) => setFont(e.target.value)}>
                {FONT_LIST.map((f) => (
                  <option key={f} value={f}>
                    {f}
                  </option>
                ))}
              </select>
            </div>
            <div className="bl-listfield">
              <span>Yazı tipi stili:</span>
              <input value={style} readOnly />
              <select size={6} value={style} onChange={(e) => setStyle(e.target.value)}>
                {STYLES.map((s) => (
                  <option key={s}>{s}</option>
                ))}
              </select>
            </div>
            <div className="bl-listfield is-narrow">
              <span>Boyut:</span>
              <input value={size} onChange={(e) => setSize(e.target.value)} />
              <select size={6} value={size} onChange={(e) => setSize(e.target.value)}>
                {FONT_SIZES.map((s) => (
                  <option key={s} value={String(s).replace(".", ",")}>
                    {String(s).replace(".", ",")}
                  </option>
                ))}
              </select>
            </div>
          </div>
          <div className="bl-dlg-row">
            <Field label="Yazı tipi rengi:">
              <ColorField value={color} onChange={setColor} noneLabel="Otomatik" />
            </Field>
            <Field label="Alt çizgi stili:">
              <select value={uType} onChange={(e) => setUType(e.target.value)}>
                <option value="none">(yok)</option>
                {UNDERLINE_TYPES.map((u) => (
                  <option key={u.id} value={u.id}>
                    {u.label}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="Alt çizgi rengi:">
              <ColorField value={uColor} onChange={setUColor} noneLabel="Otomatik" disabled={uType === "none"} />
            </Field>
          </div>
          <FieldSet legend="Efektler">
            <div className="bl-dlg-checks">
              <Check label="Üstü çizili" v={strike} set={(v) => (setStrike(v), v && setDstrike(false))} />
              <Check label="Küçük büyük harfler" v={smallCaps} set={(v) => (setSmallCaps(v), v && setCaps(false))} />
              <Check label="Çift üstü çizili" v={dstrike} set={(v) => (setDstrike(v), v && setStrike(false))} />
              <Check label="Tümü büyük harf" v={caps} set={(v) => (setCaps(v), v && setSmallCaps(false))} />
              <Check label="Üst simge" v={sup} set={(v) => (setSup(v), v && setSub(false))} />
              <Check label="Gizli" v={hidden} set={setHidden} />
              <Check label="Alt simge" v={sub} set={(v) => (setSub(v), v && setSup(false))} />
            </div>
          </FieldSet>
        </>
      ) : (
        <FieldSet legend="Karakter Aralığı">
          <div className="bl-dlg-row">
            <Field label="Aralık:">
              <select value={spacingKind} onChange={(e) => setSpacingKind(e.target.value)}>
                <option value="normal">Normal</option>
                <option value="expanded">Genişletilmiş</option>
                <option value="condensed">Sıkıştırılmış</option>
              </select>
            </Field>
            <Field label="Değer (nk):">
              <input value={spacing} disabled={spacingKind === "normal"} onChange={(e) => setSpacing(e.target.value)} />
            </Field>
          </div>
        </FieldSet>
      )}
      <FieldSet legend="Önizleme">
        <div className="bl-font-preview">
          <span style={preview}>{font}</span>
        </div>
      </FieldSet>
    </Dialog>
  );
}

export function Check({ label, v, set, disabled }: { label: string; v: boolean; set: (v: boolean) => void; disabled?: boolean }) {
  return (
    <label className="bl-check">
      <input type="checkbox" checked={v} disabled={disabled} onChange={(e) => set(e.target.checked)} />
      {label}
    </label>
  );
}

export function ColorField({ value, onChange, noneLabel, disabled }: { value: string | null; onChange: (c: string | null) => void; noneLabel: string; disabled?: boolean }) {
  return (
    <span className="bl-colorfield">
      <select value={value ? "custom" : "none"} disabled={disabled} onChange={(e) => onChange(e.target.value === "none" ? null : (value ?? "#000000"))}>
        <option value="none">{noneLabel}</option>
        <option value="custom">Renk</option>
      </select>
      <input type="color" value={value ?? "#000000"} disabled={disabled} onChange={(e) => onChange(e.target.value.toUpperCase())} />
    </span>
  );
}
