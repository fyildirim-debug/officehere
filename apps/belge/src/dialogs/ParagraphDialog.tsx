// Paragraf iletişim kutusu: girintiler, aralıklar, satır aralığı, satır ve sayfa sonları.
import { useState } from "react";
import { Dialog, DialogButton, DialogTabs, Field, FieldSet } from "@officehere/ui";
import { useBelge } from "../state/context";
import { AutoFocus } from "./AutoFocus";
import { currentPara, setParaAttrs } from "../editor/commands";
import { effectiveProps, styleIdOf } from "../model/styles";
import type { LineRule } from "../editor/paragraphFormat";
import { Check } from "./FontDialog";

const PT_PER_CM = 72 / 2.54;
const fmtNum = (n: number, d = 2) => Number(n.toFixed(d)).toLocaleString("tr-TR");
const parse = (s: string) => {
  const n = parseFloat(s.replace(",", ".").replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
};

type LineKind = "single" | "onehalf" | "double" | "atLeast" | "exact" | "multiple";

export function ParagraphDialog({ onClose }: { onClose: () => void }) {
  const ctx = useBelge();
  const { editor, styles } = ctx;
  const p = currentPara(editor);
  const a = (p?.node.attrs ?? {}) as Record<string, unknown>;
  const st = effectiveProps(styles, p ? styleIdOf(p.node.type.name, p.node.attrs) : "Normal");
  const [tab, setTab] = useState("indent");
  const [align, setAlign] = useState((a.textAlign as string) ?? st.align ?? "left");
  const [left, setLeft] = useState(fmtNum(((a.indentLeft as number | null) ?? st.indentLeft ?? 0) / PT_PER_CM));
  const [right, setRight] = useState(fmtNum(((a.indentRight as number | null) ?? st.indentRight ?? 0) / PT_PER_CM));
  const first = (a.indentFirst as number | null) ?? 0;
  const [special, setSpecial] = useState(first > 0 ? "first" : first < 0 ? "hanging" : "none");
  const [specialVal, setSpecialVal] = useState(fmtNum(Math.abs(first || 35.4) / PT_PER_CM));
  const [before, setBefore] = useState(fmtNum((a.spaceBefore as number | null) ?? st.spaceBefore ?? 0, 1));
  const [after, setAfter] = useState(fmtNum((a.spaceAfter as number | null) ?? st.spaceAfter ?? 0, 1));
  const rule = (a.lineRule as LineRule | null) ?? "auto";
  const lh = (a.lineHeight as number | null) ?? st.line ?? 1.08;
  const initKind: LineKind = rule === "exact" ? "exact" : rule === "atLeast" ? "atLeast" : lh === 1 ? "single" : lh === 1.5 ? "onehalf" : lh === 2 ? "double" : "multiple";
  const [lineKind, setLineKind] = useState<LineKind>(initKind);
  const [lineVal, setLineVal] = useState(fmtNum(rule === "auto" ? lh : lh, 2));
  const [keepNext, setKeepNext] = useState(!!a.keepNext);
  const [keepLines, setKeepLines] = useState(!!a.keepLines);
  const [pbb, setPbb] = useState(!!a.pageBreakBefore);
  const isHeading = p?.node.type.name === "heading";
  const [outline, setOutline] = useState(String(isHeading ? (p?.node.attrs.level as number) : ((a.outlineLevel as number | null) ?? 0)));

  const values = () => {
    const sv = special === "none" ? null : (special === "first" ? 1 : -1) * parse(specialVal) * PT_PER_CM;
    let lineRule: LineRule = "auto";
    let lineHeight = 1;
    switch (lineKind) {
      case "single":
        lineHeight = 1;
        break;
      case "onehalf":
        lineHeight = 1.5;
        break;
      case "double":
        lineHeight = 2;
        break;
      case "multiple":
        lineHeight = parse(lineVal) || 1;
        break;
      case "exact":
        lineRule = "exact";
        lineHeight = parse(lineVal) || 12;
        break;
      case "atLeast":
        lineRule = "atLeast";
        lineHeight = parse(lineVal) || 12;
        break;
    }
    return { lineRule, lineHeight, sv };
  };

  const apply = () => {
    const { lineRule, lineHeight, sv } = values();
    const l = parse(left) * PT_PER_CM;
    const r = parse(right) * PT_PER_CM;
    editor.chain().focus().setTextAlign(align).run();
    setParaAttrs(editor, {
      indentLeft: l ? +l.toFixed(2) : null,
      indentRight: r ? +r.toFixed(2) : null,
      indentFirst: sv ? +sv.toFixed(2) : null,
      spaceBefore: parse(before),
      spaceAfter: parse(after),
      lineRule,
      lineHeight,
      keepNext,
      keepLines,
      pageBreakBefore: pbb,
      ...(isHeading ? {} : { outlineLevel: Number(outline) || null }),
    });
    onClose();
  };

  /** Varsayılan Olarak Ayarla: bu belgedeki Normal stilin paragraf biçimi. */
  const setDefault = () => {
    const { lineHeight, lineRule } = values();
    const normal = styles.Normal;
    ctx.updateSettings((s) => ({
      ...s,
      customStyles: {
        ...(s.customStyles ?? {}),
        Normal: {
          ...normal,
          props: {
            ...normal.props,
            align: align as "left",
            spaceBefore: parse(before),
            spaceAfter: parse(after),
            ...(lineRule === "auto" ? { line: lineHeight } : {}),
            indentLeft: parse(left) * PT_PER_CM || undefined,
            indentRight: parse(right) * PT_PER_CM || undefined,
          },
        },
      },
    }));
    ctx.notify("Paragraf biçimi bu belgenin varsayılanı (Normal stili) olarak ayarlandı.", "success");
    onClose();
  };

  // Önizleme: önceki/sonraki paragraflar arasında seçilen biçim.
  const pv = values();
  const previewStyle: React.CSSProperties = {
    textAlign: align as React.CSSProperties["textAlign"],
    paddingLeft: `${parse(left) * PT_PER_CM * 0.5}pt`,
    paddingRight: `${parse(right) * PT_PER_CM * 0.5}pt`,
    textIndent: pv.sv ? `${pv.sv * 0.5}pt` : undefined,
    marginTop: `${parse(before) * 0.5}pt`,
    marginBottom: `${parse(after) * 0.5}pt`,
    lineHeight: pv.lineRule === "auto" ? pv.lineHeight * 1.2 : undefined,
  };

  return (
    <Dialog
      open
      title="Paragraf"
      width={480}
      onClose={onClose}
      onSubmit={apply}
      footer={
        <>
          <DialogButton
            onClick={() => {
              apply();
              ctx.openDialog("tabs");
            }}
          >
            Sekmeler...
          </DialogButton>
          <DialogButton onClick={setDefault}>Varsayılan Olarak Ayarla</DialogButton>
          <DialogButton primary onClick={apply}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <DialogTabs
        tabs={[
          { id: "indent", label: "Girintiler ve Aralıklar" },
          { id: "breaks", label: "Satır ve Sayfa Sonları" },
        ]}
        active={tab}
        onChange={setTab}
      />
      <AutoFocus deps={[tab]} />
      {tab === "indent" ? (
        <>
          <FieldSet legend="Genel">
            <Field label="Hizalama:" inline>
              <select value={align} onChange={(e) => setAlign(e.target.value)}>
                <option value="left">Sol</option>
                <option value="center">Ortalanmış</option>
                <option value="right">Sağ</option>
                <option value="justify">İki Yana Yaslanmış</option>
              </select>
            </Field>
            <Field label="Anahat düzeyi:" inline>
              <select aria-label="Anahat düzeyi" value={outline} disabled={isHeading} onChange={(e) => setOutline(e.target.value)}>
                <option value="0">Gövde Metni</option>
                {[1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => (
                  <option key={n} value={String(n)}>
                    Düzey {n}
                  </option>
                ))}
              </select>
            </Field>
          </FieldSet>
          <FieldSet legend="Girinti">
            <div className="bl-dlg-grid">
              <Field label="Sol (cm):">
                <input value={left} onChange={(e) => setLeft(e.target.value)} />
              </Field>
              <Field label="Özel:">
                <select value={special} onChange={(e) => setSpecial(e.target.value)}>
                  <option value="none">(yok)</option>
                  <option value="first">İlk satır</option>
                  <option value="hanging">Asılı</option>
                </select>
              </Field>
              <Field label="Sağ (cm):">
                <input value={right} onChange={(e) => setRight(e.target.value)} />
              </Field>
              <Field label="Değer (cm):">
                <input value={specialVal} disabled={special === "none"} onChange={(e) => setSpecialVal(e.target.value)} />
              </Field>
            </div>
          </FieldSet>
          <FieldSet legend="Aralık">
            <div className="bl-dlg-grid">
              <Field label="Önce (nk):">
                <input value={before} onChange={(e) => setBefore(e.target.value)} />
              </Field>
              <Field label="Satır aralığı:">
                <select value={lineKind} onChange={(e) => setLineKind(e.target.value as LineKind)}>
                  <option value="single">Tek</option>
                  <option value="onehalf">1,5 satır</option>
                  <option value="double">Çift</option>
                  <option value="atLeast">En az</option>
                  <option value="exact">Tam</option>
                  <option value="multiple">Birden çok</option>
                </select>
              </Field>
              <Field label="Sonra (nk):">
                <input value={after} onChange={(e) => setAfter(e.target.value)} />
              </Field>
              <Field label={lineKind === "exact" || lineKind === "atLeast" ? "Aralık (nk):" : "Aralık:"}>
                <input value={lineVal} disabled={lineKind === "single" || lineKind === "onehalf" || lineKind === "double"} onChange={(e) => setLineVal(e.target.value)} />
              </Field>
            </div>
          </FieldSet>
          <FieldSet legend="Önizleme">
            <div className="bl-para-preview" aria-label="Önizleme">
              <span className="bl-style-prev-dim">Önceki Paragraf Önceki Paragraf Önceki Paragraf Önceki Paragraf Önceki Paragraf</span>
              <div style={previewStyle}>Örnek Metin Örnek Metin Örnek Metin Örnek Metin Örnek Metin Örnek Metin Örnek Metin Örnek Metin</div>
              <span className="bl-style-prev-dim">Sonraki Paragraf Sonraki Paragraf Sonraki Paragraf Sonraki Paragraf</span>
            </div>
          </FieldSet>
        </>
      ) : (
        <FieldSet legend="Sayfalandırma">
          <div className="bl-dlg-checks is-single">
            <Check label="Sonrakiyle birlikte tut" v={keepNext} set={setKeepNext} />
            <Check label="Satırları birlikte tut" v={keepLines} set={setKeepLines} />
            <Check label="Önce sayfa sonu" v={pbb} set={setPbb} />
          </div>
        </FieldSet>
      )}
    </Dialog>
  );
}
