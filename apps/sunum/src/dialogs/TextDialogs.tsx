/** Yazı Tipi ve Paragraf iletişim kutuları. */
import { useState } from "react";
import { Dialog, DialogButton, DialogTabs, Field } from "@officehere/ui";
import { FONT_LIST, fontStack } from "@officehere/core";
import { cssColor } from "../model/color";
import { resolveFont } from "../model/textStyle";
import type { Color, ParaProps, RunProps } from "../model/types";
import type { Actions } from "../state/actions";
import { applyRun, common, setParaProps, type TextInfo } from "../state/textActions";
import { PaneColorButton } from "../panes/paneControls";
import { cm, fromCm } from "../ribbon/common";

export function FontDialog({ open, onClose, actions: a, info }: { open: boolean; onClose: () => void; actions: Actions; info: TextInfo | null }) {
  const runs = info?.runs ?? [];
  const theme = a.pres.theme;
  const [tab, setTab] = useState("font");
  const init = (): RunProps => ({
    font: common<string>(runs, "font"),
    size: common<number>(runs, "size"),
    bold: common<boolean>(runs, "bold"),
    italic: common<boolean>(runs, "italic"),
    underline: common<boolean>(runs, "underline"),
    strike: common<boolean>(runs, "strike"),
    color: common<Color>(runs, "color"),
    caps: common<RunProps["caps"]>(runs, "caps"),
    baseline: common<number>(runs, "baseline"),
    spacing: common<number>(runs, "spacing"),
  });
  const [p, setP] = useState<RunProps>(init);
  const [lastOpen, setLastOpen] = useState(false);
  if (open !== lastOpen) {
    setLastOpen(open);
    if (open) setP(init());
  }
  const style = p.bold && p.italic ? "bi" : p.bold ? "b" : p.italic ? "i" : "n";
  const apply = () => {
    applyRun(a, (r) => ({
      ...r,
      font: p.font ?? r.font,
      size: p.size ?? r.size,
      bold: p.bold,
      italic: p.italic,
      underline: p.underline,
      strike: p.strike,
      color: p.color ?? r.color,
      caps: p.caps,
      baseline: p.baseline,
      spacing: p.spacing,
    }));
    onClose();
  };
  return (
    <Dialog open={open} title="Yazı Tipi" onClose={onClose} onSubmit={apply} width={520} footer={<><DialogButton primary onClick={apply}>Tamam</DialogButton><DialogButton onClick={onClose}>İptal</DialogButton></>}>
      <DialogTabs tabs={[{ id: "font", label: "Yazı Tipi" }, { id: "spacing", label: "Karakter Aralığı" }]} active={tab} onChange={setTab} />
      {tab === "font" ? (
        <>
          <div className="dlg-row">
            <Field label="Latin metin yazı tipi:">
              <select value={p.font ? resolveFont(p.font, theme) : ""} onChange={(e) => setP({ ...p, font: e.target.value })} style={{ width: 200 }}>
                <option value="">(karışık)</option>
                {[theme.fonts.major, theme.fonts.minor, ...FONT_LIST.filter((f) => f !== theme.fonts.major && f !== theme.fonts.minor)].map((f) => (
                  <option key={f} value={f}>
                    {f}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="Yazı tipi stili:">
              <select
                value={style}
                onChange={(e) => {
                  const v = e.target.value;
                  setP({ ...p, bold: v === "b" || v === "bi", italic: v === "i" || v === "bi" });
                }}
              >
                <option value="n">Normal</option>
                <option value="i">İtalik</option>
                <option value="b">Kalın</option>
                <option value="bi">Kalın İtalik</option>
              </select>
            </Field>
            <Field label="Boyut:">
              <input type="number" min={1} max={4000} step={0.5} value={p.size ?? ""} onChange={(e) => setP({ ...p, size: e.target.value ? Number(e.target.value) : undefined })} style={{ width: 70 }} />
            </Field>
          </div>
          <div className="dlg-row">
            <span>Yazı tipi rengi:</span>
            <PaneColorButton theme={theme} value={p.color} onChange={(c) => setP({ ...p, color: c ?? undefined })} />
            <label>
              <input type="checkbox" checked={!!p.underline} onChange={(e) => setP({ ...p, underline: e.target.checked })} /> Altı çizili
            </label>
          </div>
          <fieldset className="oh-fieldset">
            <legend>Efektler</legend>
            <div className="dlg-row">
              <label>
                <input type="checkbox" checked={!!p.strike} onChange={(e) => setP({ ...p, strike: e.target.checked })} /> Üstü çizili
              </label>
              <label>
                <input type="checkbox" checked={(p.baseline ?? 0) > 0} onChange={(e) => setP({ ...p, baseline: e.target.checked ? 30 : undefined })} /> Üst simge
              </label>
              <label>
                <input type="checkbox" checked={(p.baseline ?? 0) < 0} onChange={(e) => setP({ ...p, baseline: e.target.checked ? -25 : undefined })} /> Alt simge
              </label>
            </div>
            <div className="dlg-row">
              <label>
                <input type="checkbox" checked={p.caps === "small"} onChange={(e) => setP({ ...p, caps: e.target.checked ? "small" : undefined })} /> Küçük büyük harf
              </label>
              <label>
                <input type="checkbox" checked={p.caps === "all"} onChange={(e) => setP({ ...p, caps: e.target.checked ? "all" : undefined })} /> Tümü büyük harf
              </label>
            </div>
          </fieldset>
        </>
      ) : (
        <div className="dlg-row">
          <Field label="Aralık:">
            <select
              value={p.spacing === undefined ? "normal" : p.spacing > 0 ? "exp" : "cond"}
              onChange={(e) => setP({ ...p, spacing: e.target.value === "normal" ? undefined : e.target.value === "exp" ? Math.abs(p.spacing ?? 1) : -Math.abs(p.spacing ?? 1) })}
            >
              <option value="normal">Normal</option>
              <option value="exp">Genişletilmiş</option>
              <option value="cond">Daraltılmış</option>
            </select>
          </Field>
          <Field label="Uzunluk (nk):">
            <input type="number" step={0.1} value={Math.abs(p.spacing ?? 0)} onChange={(e) => setP({ ...p, spacing: p.spacing === undefined ? undefined : Math.sign(p.spacing || 1) * Number(e.target.value) })} style={{ width: 70 }} />
          </Field>
        </div>
      )}
      <div className="dlg-preview" style={{ fontFamily: fontStack(resolveFont(p.font, theme)), fontWeight: p.bold ? 700 : 400, fontStyle: p.italic ? "italic" : undefined, textDecoration: [p.underline ? "underline" : "", p.strike ? "line-through" : ""].join(" ").trim() || undefined, color: cssColor(p.color, theme), letterSpacing: p.spacing ? `${p.spacing}pt` : undefined, fontVariant: p.caps === "small" ? "small-caps" : undefined, textTransform: p.caps === "all" ? "uppercase" : undefined, fontSize: Math.min(28, (p.size ?? 18) * 1.1) }}>
        Örnek Metin AaBbÇçĞğİı
      </div>
    </Dialog>
  );
}

export function ParagraphDialog({ open, onClose, actions: a, info }: { open: boolean; onClose: () => void; actions: Actions; info: TextInfo | null }) {
  const init = (): ParaProps & { special: "none" | "first" | "hanging"; by: number } => {
    const p = info?.para ?? {};
    const ind = p.indent ?? 0;
    return { ...p, special: ind > 0 ? "first" : ind < 0 ? "hanging" : "none", by: Math.abs(ind) };
  };
  const [p, setP] = useState(init);
  const [lastOpen, setLastOpen] = useState(false);
  if (open !== lastOpen) {
    setLastOpen(open);
    if (open) setP(init());
  }
  const lsMode = p.lineSpacingPt !== undefined ? "exact" : (p.lineSpacing ?? 100) === 100 ? "single" : (p.lineSpacing ?? 100) === 150 ? "1.5" : (p.lineSpacing ?? 100) === 200 ? "double" : "multiple";
  const apply = () => {
    const indent = p.special === "none" ? 0 : p.special === "first" ? p.by : -p.by;
    const textStart = p.special === "hanging" ? (p.marL ?? 0) : (p.marL ?? 0);
    setParaProps(a, { align: p.align, marL: textStart, indent, lineSpacing: p.lineSpacingPt !== undefined ? undefined : p.lineSpacing, lineSpacingPt: p.lineSpacingPt, spaceBefore: p.spaceBefore, spaceAfter: p.spaceAfter });
    onClose();
  };
  return (
    <Dialog open={open} title="Paragraf" onClose={onClose} onSubmit={apply} width={520} footer={<><DialogButton primary onClick={apply}>Tamam</DialogButton><DialogButton onClick={onClose}>İptal</DialogButton></>}>
      <fieldset className="oh-fieldset">
        <legend>Genel</legend>
        <div className="dlg-row">
          <label>
            Hizalama:
            <select value={p.align ?? "l"} onChange={(e) => setP({ ...p, align: e.target.value as ParaProps["align"] })}>
              <option value="l">Sol</option>
              <option value="ctr">Orta</option>
              <option value="r">Sağ</option>
              <option value="just">İki yana yaslanmış</option>
            </select>
          </label>
        </div>
      </fieldset>
      <fieldset className="oh-fieldset">
        <legend>Girinti</legend>
        <div className="dlg-row">
          <label>
            Metinden önce:
            <input type="number" step={0.1} min={0} value={cm(p.marL ?? 0)} onChange={(e) => setP({ ...p, marL: fromCm(Number(e.target.value)) })} style={{ width: 70 }} /> cm
          </label>
          <label>
            Özel:
            <select value={p.special} onChange={(e) => setP({ ...p, special: e.target.value as "none" | "first" | "hanging" })}>
              <option value="none">(yok)</option>
              <option value="first">İlk Satır</option>
              <option value="hanging">Asılı</option>
            </select>
          </label>
          <label>
            Ölçü:
            <input type="number" step={0.1} min={0} value={cm(p.by)} onChange={(e) => setP({ ...p, by: fromCm(Number(e.target.value)) })} style={{ width: 70 }} /> cm
          </label>
        </div>
      </fieldset>
      <fieldset className="oh-fieldset">
        <legend>Aralık</legend>
        <div className="dlg-row">
          <label>
            Önce:
            <input type="number" min={0} step={1} value={p.spaceBefore ?? 0} onChange={(e) => setP({ ...p, spaceBefore: Number(e.target.value) })} style={{ width: 60 }} /> nk
          </label>
          <label>
            Satır Aralığı:
            <select
              value={lsMode}
              onChange={(e) => {
                const v = e.target.value;
                if (v === "single") setP({ ...p, lineSpacing: 100, lineSpacingPt: undefined });
                else if (v === "1.5") setP({ ...p, lineSpacing: 150, lineSpacingPt: undefined });
                else if (v === "double") setP({ ...p, lineSpacing: 200, lineSpacingPt: undefined });
                else if (v === "exact") setP({ ...p, lineSpacingPt: 24 });
                else setP({ ...p, lineSpacing: 120, lineSpacingPt: undefined });
              }}
            >
              <option value="single">Tek</option>
              <option value="1.5">1,5 satır</option>
              <option value="double">Çift</option>
              <option value="exact">Tam olarak</option>
              <option value="multiple">Birden çok</option>
            </select>
          </label>
        </div>
        <div className="dlg-row">
          <label>
            Sonra:
            <input type="number" min={0} step={1} value={p.spaceAfter ?? 0} onChange={(e) => setP({ ...p, spaceAfter: Number(e.target.value) })} style={{ width: 60 }} /> nk
          </label>
          <label>
            Aralık:
            {lsMode === "exact" ? (
              <input type="number" min={1} step={1} value={p.lineSpacingPt ?? 24} onChange={(e) => setP({ ...p, lineSpacingPt: Number(e.target.value) })} style={{ width: 60 }} />
            ) : (
              <input type="number" min={0.5} step={0.1} value={(p.lineSpacing ?? 100) / 100} onChange={(e) => setP({ ...p, lineSpacing: Number(e.target.value) * 100 })} style={{ width: 60 }} />
            )}
          </label>
        </div>
      </fieldset>
    </Dialog>
  );
}
