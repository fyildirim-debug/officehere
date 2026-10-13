// Büyük İlk Harf iletişim kutusu: konum (yok / metin içinde / kenar boşluğunda), yazı tipi, satır sayısı, uzaklık.
import { useMemo, useState } from "react";
import { FONT_LIST, fontStack } from "@officehere/core";
import { Dialog, DialogButton, Field, FieldSet } from "@officehere/ui";
import { useBelge } from "../state/context";
import { DROP_CAP_DEFAULT, canDropCap, dropCapTarget, setDropCap, type DropCap, type DropCapMode } from "../editor/dropCap";

const PT_PER_CM = 72 / 2.54;

function Preview({ mode }: { mode: DropCapMode | "none" }) {
  return (
    <span className={`bl-dc-preview is-${mode}`} aria-hidden="true">
      {mode !== "none" && <b>A</b>}
      <i />
      <i />
      <i />
      <i />
    </span>
  );
}

export function DropCapDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const target = useMemo(() => dropCapTarget(editor.state), [editor]);
  const cur = (target?.node.attrs.dropCap as DropCap | null) ?? null;
  const [mode, setMode] = useState<DropCapMode | "none">(cur?.mode ?? "drop");
  const [font, setFont] = useState(cur?.font ?? "");
  const [lines, setLines] = useState(String(cur?.lines ?? DROP_CAP_DEFAULT.lines));
  const [dist, setDist] = useState(((cur?.distance ?? 0) / PT_PER_CM).toFixed(2).replace(".", ","));
  const allowed = !!target && canDropCap(target.node);
  const n = Math.round(Number(lines));
  const cm = parseFloat(dist.replace(",", "."));
  const valid = mode === "none" || (n >= 1 && n <= 10 && Number.isFinite(cm) && cm >= 0 && cm <= 5);
  const apply = () => {
    if (!valid) return;
    if (mode === "none") setDropCap(editor, null);
    else setDropCap(editor, { mode, lines: n, distance: +(cm * PT_PER_CM).toFixed(2), font: font || null });
    onClose();
  };
  const radio = (m: DropCapMode | "none", label: string) => (
    <label className={`bl-dc-pos${mode === m ? " is-selected" : ""}`}>
      <Preview mode={m} />
      <span>
        <input type="radio" name="dcpos" checked={mode === m} disabled={m !== "none" && !allowed} onChange={() => setMode(m)} /> {label}
      </span>
    </label>
  );
  return (
    <Dialog
      open
      title="Büyük İlk Harf"
      width={420}
      onClose={onClose}
      onSubmit={apply}
      footer={
        <>
          <DialogButton primary onClick={apply} disabled={!target || !valid}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <FieldSet legend="Konum">
        <div className="bl-dc-positions">
          {radio("none", "Yok")}
          {radio("drop", "Metin İçinde")}
          {radio("margin", "Kenar Boşluğunda")}
        </div>
      </FieldSet>
      <FieldSet legend="Seçenekler">
        <Field label="Yazı tipi:">
          <select value={font} disabled={mode === "none"} onChange={(e) => setFont(e.target.value)} aria-label="Yazı tipi">
            <option value="">(Paragrafın yazı tipi)</option>
            {FONT_LIST.map((f) => (
              <option key={f} value={f} style={{ fontFamily: fontStack(f) }}>
                {f}
              </option>
            ))}
          </select>
        </Field>
        <div className="bl-dlg-grid">
          <Field label="Bırakılacak satır:">
            <input type="number" min={1} max={10} value={lines} disabled={mode === "none"} onChange={(e) => setLines(e.target.value)} aria-label="Bırakılacak satır" />
          </Field>
          <Field label="Metinden uzaklık (cm):">
            <input value={dist} disabled={mode === "none"} onChange={(e) => setDist(e.target.value)} aria-label="Metinden uzaklık" />
          </Field>
        </div>
      </FieldSet>
      {!allowed && <p className="bl-dlg-note">Büyük ilk harf için imleç metinle başlayan bir paragrafta olmalıdır.</p>}
    </Dialog>
  );
}
