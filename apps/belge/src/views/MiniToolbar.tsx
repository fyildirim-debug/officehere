// Mini araç çubuğu: fareyle metin seçilince seçimin üstünde, sağ tıkta bağlam menüsünün üstünde açılan
// küçük biçimlendirme çubuğu (yazı tipi, boyut, büyült/küçült, stiller, kalın/italik/altı çizili, vurgu,
// yazı rengi, madde işaretleri, numaralandırma, açıklama).
import { syncDomSelection } from "../editor/pagination";
import { useEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import {
  CommentAdd20Regular,
  Highlight20Regular,
  PaintBrush20Regular,
  TextBold20Regular,
  TextBulletListLtr20Regular,
  TextColor20Regular,
  TextItalic20Regular,
  TextNumberListLtr20Regular,
  TextUnderline20Regular,
  TextEditStyle20Regular,
} from "@fluentui/react-icons";
import { FONT_LIST, FONT_SIZES, fontStack } from "@officehere/core";
import { ColorSplitButton, Menu, MenuItem, RibbonButton, RibbonCombo, RibbonMenuButton } from "@officehere/ui";
import { useBelge } from "../state/context";
import { applyStyle, setFontSize, stepFontSize } from "../editor/commands";
import { applyBulletMain, applyNumberMain } from "../ribbon/lists";

export interface MiniBarAnchor {
  x: number;
  y: number;
  /** Bağlam menüsüyle birlikte mi açıldı. */
  context: boolean;
}

const QUICK_STYLES = [
  ["Normal", "Normal"],
  ["NoSpacing", "Aralık Yok"],
  ["Heading1", "Başlık 1"],
  ["Heading2", "Başlık 2"],
  ["Title", "Konu Başlığı"],
  ["Quote", "Alıntı"],
] as const;

/** Seçimin üstünde konumlanan çubuk (belge düzenleyicisinin fare seçimini izler). */
export function useMiniToolbar() {
  const ctx = useBelge();
  const [anchor, setAnchor] = useState<MiniBarAnchor | null>(null);
  const ed = ctx.editor;

  useEffect(() => {
    const dom = ed.view.dom as HTMLElement;
    let down = false;
    const onDown = (e: MouseEvent) => {
      down = e.button === 0;
      if (!(e.target as HTMLElement).closest?.(".bl-minibar")) setAnchor((a) => (a?.context ? a : null));
    };
    const onUp = (e: MouseEvent) => {
      if (!down || e.button !== 0) return;
      down = false;
      window.setTimeout(() => {
        if (ed.isDestroyed) return;
        // Çift/üçlü tıklamada seçim DOM'da değişmiş ama düzenleyiciye henüz okunmamış olabilir.
        if (ed.view.hasFocus()) syncDomSelection(ed.view);
        const sel = ed.state.selection;
        if (sel.empty || !ed.view.hasFocus() || !ed.state.doc.textBetween(sel.from, sel.to).trim()) return;
        const r = window.getSelection()?.getRangeAt(0)?.getBoundingClientRect();
        if (!r || (r.width === 0 && r.height === 0)) return;
        setAnchor({ x: Math.max(8, r.left), y: r.top, context: false });
      }, 0);
    };
    const onKey = () => setAnchor(null);
    dom.addEventListener("mousedown", onDown);
    window.addEventListener("mouseup", onUp);
    dom.addEventListener("keydown", onKey);
    const onGlobalDown = (e: MouseEvent) => {
      const t = e.target as HTMLElement;
      if (!t.closest?.(".bl-minibar") && !t.closest?.("[data-oh-popover]") && !dom.contains(t)) setAnchor(null);
    };
    document.addEventListener("mousedown", onGlobalDown, true);
    // Yalnız belge kaydırılınca kapan; çubuğun kendi açılır listeleri (yazı tipi, boyut) kaydırılınca açık kalsın.
    const onScroll = (e: Event) => {
      const t = e.target;
      if (t instanceof Element && (t.closest("[data-oh-popover]") || t.closest(".bl-minibar"))) return;
      setAnchor((a) => (a && !a.context ? null : a));
    };
    document.addEventListener("scroll", onScroll, true);
    // Bağlam menüsü kapanınca (Esc ya da komut) çubuk da kapanır.
    const afterInput = (e: Event) => {
      if (e instanceof KeyboardEvent && e.key === "Escape") return setAnchor(null);
      window.setTimeout(() => setAnchor((a) => (a?.context && !document.querySelector("[data-oh-popover]") ? null : a)), 60);
    };
    document.addEventListener("keyup", afterInput, true);
    document.addEventListener("mouseup", afterInput, true);
    return () => {
      dom.removeEventListener("mousedown", onDown);
      window.removeEventListener("mouseup", onUp);
      dom.removeEventListener("keydown", onKey);
      document.removeEventListener("mousedown", onGlobalDown, true);
      document.removeEventListener("scroll", onScroll, true);
      document.removeEventListener("keyup", afterInput, true);
      document.removeEventListener("mouseup", afterInput, true);
    };
  }, [ed]);

  return { anchor, setAnchor };
}

export function MiniToolbar({ anchor, onClose }: { anchor: MiniBarAnchor | null; onClose: () => void }) {
  const ctx = useBelge();
  const { editor, fmt, styles } = ctx;
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null);
  const [hl, setHl] = useState<string | null>("#FFFF00");
  const [color, setColor] = useState<string | null>("#C00000");

  useEffect(() => {
    if (!anchor) return setPos(null);
    const el = ref.current;
    const w = el?.offsetWidth ?? 360;
    const h = el?.offsetHeight ?? 70;
    const left = Math.min(Math.max(4, anchor.x - (anchor.context ? 0 : 20)), window.innerWidth - w - 4);
    const top = anchor.context ? Math.max(4, anchor.y - h - 6) : anchor.y - h - 10 < 4 ? anchor.y + 28 : anchor.y - h - 10;
    setPos({ left, top });
  }, [anchor]);

  if (!anchor) return null;
  const chain = () => editor.chain().focus();
  const run = (f: () => void) => () => {
    f();
    if (anchor.context) onClose();
  };
  return createPortal(
    <div ref={ref} className="bl-minibar" role="toolbar" aria-label="Mini araç çubuğu" style={pos ? { left: pos.left, top: pos.top } : { left: -9999, top: 0 }} onMouseDown={(e) => e.target instanceof HTMLInputElement || e.preventDefault()}>
      <div className="bl-minibar-row">
        <RibbonCombo value={fmt.font} width={120} options={FONT_LIST.map((f) => ({ value: f, style: { fontFamily: fontStack(f), fontSize: 14 } }))} onCommit={(v) => chain().setFontFamily(v).run()} tip="Yazı Tipi" />
        <RibbonCombo
          value={String(fmt.size).replace(".", ",")}
          width={46}
          options={FONT_SIZES.map((s) => String(s).replace(".", ","))}
          onCommit={(v) => {
            const n = parseFloat(v.replace(",", "."));
            if (Number.isFinite(n) && n > 0) setFontSize(editor, n);
          }}
          tip="Yazı Tipi Boyutu"
        />
        <RibbonButton size="small" icon={<span className="bl-grow">A<sup>▲</sup></span>} tip="Yazı Tipini Büyült" onClick={() => stepFontSize(editor, styles, 1)} />
        <RibbonButton size="small" icon={<span className="bl-grow is-shrink">A<sup>▼</sup></span>} tip="Yazı Tipini Küçült" onClick={() => stepFontSize(editor, styles, -1)} />
        <RibbonButton size="small" icon={<PaintBrush20Regular />} tip="Biçim Boyacısı" checked={ctx.formatPainter.active} onClick={run(() => ctx.startFormatPainter(false))} />
        <RibbonMenuButton
          size="small"
          icon={<TextEditStyle20Regular />}
          tip="Stiller"
          menu={
            <Menu width={180}>
              {QUICK_STYLES.map(([id, label]) => (
                <MenuItem key={id} label={label} checked={fmt.styleId === id} onClick={() => applyStyle(editor, styles, id)} />
              ))}
            </Menu>
          }
        />
      </div>
      <div className="bl-minibar-row">
        <RibbonButton size="small" icon={<TextBold20Regular />} tip="Kalın" shortcut="Ctrl+B" checked={fmt.bold} onClick={run(() => chain().toggleBold().run())} />
        <RibbonButton size="small" icon={<TextItalic20Regular />} tip="İtalik" shortcut="Ctrl+I" checked={fmt.italic} onClick={run(() => chain().toggleItalic().run())} />
        <RibbonButton size="small" icon={<TextUnderline20Regular />} tip="Altı Çizili" shortcut="Ctrl+U" checked={!!fmt.underline} onClick={run(() => chain().toggleUnderline().run())} />
        <ColorSplitButton icon={<Highlight20Regular />} tip="Metin Vurgu Rengi" color={hl} variant="highlight" noneLabel="Renk Yok" onColorChange={setHl} onApply={(c) => (c ? chain().setHighlight({ color: c }).run() : chain().unsetHighlight().run())} />
        <ColorSplitButton icon={<TextColor20Regular />} tip="Yazı Tipi Rengi" color={color} noneLabel="Otomatik" onColorChange={setColor} onApply={(c) => (c ? chain().setColor(c).run() : chain().unsetColor().run())} />
        <RibbonButton size="small" icon={<TextBulletListLtr20Regular />} tip="Madde İşaretleri" checked={fmt.list === "bullet"} onClick={run(() => applyBulletMain(editor))} />
        <RibbonButton size="small" icon={<TextNumberListLtr20Regular />} tip="Numaralandırma" checked={fmt.list === "ordered"} onClick={run(() => applyNumberMain(editor))} />
        {ctx.inBody && <RibbonButton size="small" icon={<CommentAdd20Regular />} tip="Yeni Açıklama" onClick={run(ctx.newComment)} />}
      </div>
    </div>,
    document.body,
  );
}

/** Sağlayıcı içinde çizilen mini çubuk; bağlam menüsü `controlRef` ile açar. */
export function MiniToolbarHost({ controlRef }: { controlRef: { current: ((a: MiniBarAnchor | null) => void) | null } }) {
  const { anchor, setAnchor } = useMiniToolbar();
  controlRef.current = setAnchor;
  return <MiniToolbar anchor={anchor} onClose={() => setAnchor(null)} />;
}
