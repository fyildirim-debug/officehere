// Resim Kırp (Resim Biçimi › Kırp): seçili resmin kırpılmamış hâli soluk gösterilir; kırpma çerçevesinin
// kenar/köşe tutamaçları sürüklenerek kırpılır. Enter, Esc ya da dışarı tıklama kırpmayı bitirir.
import { useEffect, useState, useSyncExternalStore } from "react";
import type { Editor } from "@tiptap/core";
import { NodeSelection } from "@tiptap/pm/state";
import { croppedSize, type ImageCrop } from "../editor/image";

/* Kırpma kipi (şerit düğmesi ile katman arasında paylaşılır). */
let cropPos: number | null = null;
const listeners = new Set<() => void>();
export function setCropMode(pos: number | null) {
  cropPos = pos;
  listeners.forEach((l) => l());
}
export const cropMode = () => cropPos;
const subscribe = (l: () => void) => {
  listeners.add(l);
  return () => listeners.delete(l);
};

const NONE: ImageCrop = { l: 0, t: 0, r: 0, b: 0 };
type Handle = "l" | "r" | "t" | "b" | "tl" | "tr" | "bl" | "br";
const HANDLES: Handle[] = ["tl", "t", "tr", "r", "br", "b", "bl", "l"];

/** Kırpmayı resme uygular (görüntülenen boyut ölçek korunarak değişir). */
export function applyCrop(editor: Editor, pos: number, crop: ImageCrop | null) {
  const node = editor.state.doc.nodeAt(pos);
  if (!node || node.type.name !== "image") return;
  const clean = crop && (crop.l || crop.t || crop.r || crop.b) ? crop : null;
  const size = croppedSize(node.attrs.width as number, node.attrs.height as number, node.attrs.crop as ImageCrop | null, clean);
  const tr = editor.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, crop: clean, ...size });
  editor.view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)));
}

export function CropOverlay({ editor, pagesRef, scale }: { editor: Editor; pagesRef: React.RefObject<HTMLDivElement | null>; scale: number }) {
  const pos = useSyncExternalStore(subscribe, cropMode);
  const [crop, setCrop] = useState<ImageCrop>(NONE);
  const [box, setBox] = useState<{ x: number; y: number; w: number; h: number; src: string } | null>(null);

  // Kipe girince resmin tam (kırpılmamış) kutusu hesaplanır.
  useEffect(() => {
    if (pos === null) return setBox(null);
    const node = editor.state.doc.nodeAt(pos);
    const img = (editor.view.nodeDOM(pos) as HTMLElement | null)?.querySelector("img");
    const host = pagesRef.current;
    if (!node || node.type.name !== "image" || !img || !host) return setCropMode(null);
    const c = (node.attrs.crop as ImageCrop | null) ?? NONE;
    const r = img.getBoundingClientRect();
    const hr = host.getBoundingClientRect();
    const w = r.width / scale / Math.max(0.01, 1 - c.l - c.r);
    const h = r.height / scale / Math.max(0.01, 1 - c.t - c.b);
    setCrop(c);
    setBox({ x: (r.left - hr.left) / scale - c.l * w, y: (r.top - hr.top) / scale - c.t * h, w, h, src: String(node.attrs.src) });
    // Kırpma sırasında belgedeki resim gizlenir; katman kırpılan bölgeleri soluk gösterir.
    img.style.visibility = "hidden";
    return () => {
      img.style.visibility = "";
    };
  }, [pos, editor, pagesRef, scale]);

  // Bitirme: Enter/Esc, belge değişimi ya da dışarı tıklama.
  useEffect(() => {
    if (pos === null) return;
    const finish = (commit: boolean) => {
      if (commit) applyCrop(editor, pos, crop);
      setCropMode(null);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Enter" || e.key === "Escape") {
        e.preventDefault();
        e.stopPropagation();
        finish(true);
      }
    };
    const onDown = (e: MouseEvent) => {
      if ((e.target as HTMLElement).closest(".bl-crop, .oh-ribbon")) return;
      finish(true);
    };
    window.addEventListener("keydown", onKey, true);
    window.addEventListener("mousedown", onDown, true);
    return () => {
      window.removeEventListener("keydown", onKey, true);
      window.removeEventListener("mousedown", onDown, true);
    };
  }, [pos, crop, editor]);

  if (pos === null || !box) return null;
  const startDrag = (h: Handle) => (e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const sx = e.clientX;
    const sy = e.clientY;
    const c0 = crop;
    const move = (ev: MouseEvent) => {
      const dx = (ev.clientX - sx) / scale / box.w;
      const dy = (ev.clientY - sy) / scale / box.h;
      const n = { ...c0 };
      const lim = (v: number, other: number) => Math.max(0, Math.min(0.95 - other, v));
      if (h === "l" || h === "tl" || h === "bl") n.l = lim(c0.l + dx, c0.r);
      if (h === "r" || h === "tr" || h === "br") n.r = lim(c0.r - dx, c0.l);
      if (h === "t" || h === "tl" || h === "tr") n.t = lim(c0.t + dy, c0.b);
      if (h === "b" || h === "bl" || h === "br") n.b = lim(c0.b - dy, c0.t);
      setCrop(n);
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  };
  const cx = box.x + crop.l * box.w;
  const cy = box.y + crop.t * box.h;
  const cw = box.w * (1 - crop.l - crop.r);
  const ch = box.h * (1 - crop.t - crop.b);
  return (
    <div className="bl-crop" aria-label="Kırpma">
      <img className="bl-crop-ghost" src={box.src} alt="" style={{ left: box.x, top: box.y, width: box.w, height: box.h }} draggable={false} />
      <div className="bl-crop-win" style={{ left: cx, top: cy, width: cw, height: ch }}>
        <img src={box.src} alt="" style={{ left: box.x - cx, top: box.y - cy, width: box.w, height: box.h }} draggable={false} />
        {HANDLES.map((h) => (
          <span key={h} className={`bl-crop-h is-${h}`} onMouseDown={startDrag(h)} aria-label={`Kırpma tutamacı ${h}`} />
        ))}
      </div>
    </div>
  );
}
