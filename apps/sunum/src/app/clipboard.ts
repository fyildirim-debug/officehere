/** Sistem panosu: şekil/slayt kopyala-kes-yapıştır, resim ve düz metin yapıştırma. */
import { useEffect } from "react";
import { plainText } from "../model/textStyle";
import { flatShapes } from "../model/ops";
import type { Actions } from "../state/actions";
import type { ClipboardData } from "../state/store";

export const SHAPES_MIME = "application/x-officehere-sunum";

export function readImageFile(file: Blob): Promise<{ src: string; w: number; h: number }> {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => {
      const src = String(fr.result);
      const img = new Image();
      img.onload = () => resolve({ src, w: img.naturalWidth || 400, h: img.naturalHeight || 300 });
      img.onerror = () => reject(new Error("Resim okunamadı"));
      img.src = src;
    };
    fr.onerror = () => reject(fr.error);
    fr.readAsDataURL(file);
  });
}

export async function insertImageFiles(actions: Actions, files: File[]) {
  for (const f of files) {
    if (!f.type.startsWith("image/")) continue;
    try {
      const img = await readImageFile(f);
      actions.insertPicture(img.src, img.w, img.h, "Resim");
    } catch {
      actions.message("Resim eklenemedi.", "error");
    }
  }
}

function isField(t: EventTarget | null) {
  const el = t as HTMLElement | null;
  return !!el && (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.isContentEditable);
}

export function useClipboard(actions: Actions, blocked: () => boolean) {
  useEffect(() => {
    const doCopy = (e: ClipboardEvent, cut: boolean) => {
      if (blocked() || isField(e.target) || actions.s.editing) return;
      const s = actions.s;
      if (s.focus === "slides" || (!s.sel.length && s.pres.slides.length)) {
        actions.copySlides(cut);
      } else if (!actions.copySelection(cut)) return;
      const cb = actions.store.clipboard;
      if (!cb) return;
      e.preventDefault();
      const text = (cb.shapes ?? []).flatMap((x) => flatShapes([x])).map((x) => (x.kind === "sp" ? plainText(x.text) : "")).filter(Boolean).join("\n");
      try {
        e.clipboardData?.setData(SHAPES_MIME, JSON.stringify(cb));
        e.clipboardData?.setData("text/plain", text);
      } catch {
        // yok say
      }
    };
    const onCopy = (e: ClipboardEvent) => doCopy(e, false);
    const onCut = (e: ClipboardEvent) => doCopy(e, true);
    const onPaste = (e: ClipboardEvent) => {
      if (blocked() || isField(e.target) || actions.s.editing) return;
      const dt = e.clipboardData;
      if (!dt) return;
      e.preventDefault();
      const json = dt.getData(SHAPES_MIME);
      if (json) {
        try {
          actions.store.clipboard = JSON.parse(json) as ClipboardData;
          actions.pasteInternal();
          return;
        } catch {
          // devam
        }
      }
      const files = Array.from(dt.files ?? []).filter((f) => f.type.startsWith("image/"));
      if (files.length) {
        void insertImageFiles(actions, files);
        return;
      }
      const text = dt.getData("text/plain");
      if (text) {
        actions.pasteText(text);
        return;
      }
      actions.pasteInternal();
    };
    document.addEventListener("copy", onCopy);
    document.addEventListener("cut", onCut);
    document.addEventListener("paste", onPaste);
    return () => {
      document.removeEventListener("copy", onCopy);
      document.removeEventListener("cut", onCut);
      document.removeEventListener("paste", onPaste);
    };
  }, [actions, blocked]);
}

/** Şerit "Yapıştır" düğmesi: önce sistem panosu, yoksa iç pano. */
export async function pasteFromRibbon(actions: Actions) {
  const h = actions.textHandle;
  if (actions.s.editing && h) {
    const paras = actions.store.clipboard?.paragraphs;
    const plain = paras?.map((p) => p.runs.map((r) => r.text).join("")).join("\n").replace(/\v/g, "\n");
    try {
      const t = await navigator.clipboard.readText();
      // Sistem panosu iç panoyla aynıysa biçimli paragraflar yapıştırılır.
      if (paras && t.replace(/\r\n?/g, "\n") === plain) h.insertRich(paras);
      else if (t) h.insert(t);
      return;
    } catch {
      if (paras) h.insertRich(paras);
      return;
    }
  }
  try {
    const items = await navigator.clipboard.read();
    for (const it of items) {
      const imgType = it.types.find((t) => t.startsWith("image/"));
      if (imgType && !actions.store.clipboard) {
        const blob = await it.getType(imgType);
        const img = await readImageFile(blob);
        actions.insertPicture(img.src, img.w, img.h);
        return;
      }
    }
  } catch {
    // izin yoksa iç panoya düş
  }
  if (actions.pasteInternal()) return;
  try {
    const t = await navigator.clipboard.readText();
    if (t) actions.pasteText(t);
  } catch {
    actions.message("Pano içeriğine erişilemedi. Ctrl+V kısayolunu kullanın.", "warning");
  }
}
