// Kırpma seçenekleri: kenarlardan kırpılacak yüzdeler (önizlemeli). Görüntülenen boyut ölçek korunarak değişir.
import { useState } from "react";
import { NodeSelection } from "@tiptap/pm/state";
import { Dialog, DialogButton, Field } from "@officehere/ui";
import { useBelge } from "../state/context";
import type { ImageCrop } from "../editor/image";
import { applyCrop } from "../views/CropOverlay";

export function CropDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const sel = editor.state.selection;
  const node = sel instanceof NodeSelection && sel.node.type.name === "image" ? sel.node : null;
  const c0 = (node?.attrs.crop as ImageCrop | null) ?? { l: 0, t: 0, r: 0, b: 0 };
  const [v, setV] = useState({ l: c0.l * 100, t: c0.t * 100, r: c0.r * 100, b: c0.b * 100 });
  if (!node) return null;
  const pos = sel.from;
  const clamp = (x: number) => Math.max(0, Math.min(95, Number.isFinite(x) ? x : 0));
  const crop: ImageCrop = { l: clamp(v.l) / 100, t: clamp(v.t) / 100, r: clamp(v.r) / 100, b: clamp(v.b) / 100 };
  const valid = crop.l + crop.r < 0.95 && crop.t + crop.b < 0.95;
  const apply = () => {
    if (!valid) return;
    applyCrop(editor, pos, crop);
    onClose();
  };
  const field = (k: keyof typeof v, label: string) => (
    <Field label={label}>
      <input type="number" min={0} max={95} step={1} value={+v[k].toFixed(1)} onChange={(e) => setV((x) => ({ ...x, [k]: Number(e.target.value) }))} aria-label={label} />
    </Field>
  );
  const src = String(node.attrs.src);
  return (
    <Dialog
      open
      title="Kırpma Seçenekleri"
      width={460}
      onClose={onClose}
      onSubmit={apply}
      footer={
        <>
          <DialogButton primary onClick={apply} disabled={!valid}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <div className="bl-dlg-grid">
        {field("l", "Sol (%):")}
        {field("r", "Sağ (%):")}
        {field("t", "Üst (%):")}
        {field("b", "Alt (%):")}
      </div>
      <div className="bl-crop-prev" aria-label="Önizleme">
        <img src={src} alt="" />
        <span className="bl-crop-prev-win" style={{ left: `${crop.l * 100}%`, top: `${crop.t * 100}%`, right: `${crop.r * 100}%`, bottom: `${crop.b * 100}%` }} />
      </div>
      {!valid && <div className="bl-dlg-msg">Kırpılacak alan resmin tamamını kaplayamaz.</div>}
    </Dialog>
  );
}
