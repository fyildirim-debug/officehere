// Mini araç çubuğu: fareyle metin seçilince seçimin üstünde, sağ tıkta bağlam menüsünün üstünde açılan
// küçük biçimlendirme çubuğu (yazı tipi, boyut, büyült/küçült, stiller, kalın/italik/altı çizili, vurgu,
// yazı rengi, madde işaretleri, numaralandırma, açıklama).
import { syncDomSelection } from "../editor/pagination";
import { useState } from "react";
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
import { ColorSplitButton, Menu, MenuItem, MiniToolbarHost as UiMiniToolbarHost, MiniToolbarRow, RibbonButton, RibbonCombo, RibbonMenuButton, type MiniBarAnchor } from "@officehere/ui";
import { useBelge } from "../state/context";
import { applyStyle, setFontSize, stepFontSize } from "../editor/commands";
import { applyBulletMain, applyNumberMain } from "../ribbon/lists";

export type { MiniBarAnchor };

const QUICK_STYLES = [
  ["Normal", "Normal"],
  ["NoSpacing", "Aralık Yok"],
  ["Heading1", "Başlık 1"],
  ["Heading2", "Başlık 2"],
  ["Title", "Konu Başlığı"],
  ["Quote", "Alıntı"],
] as const;

/** Çubuğun içeriği: komutlar şeritteki Giriş sekmesinin eylemlerini çağırır. */
function BelgeMiniBar({ anchor, onClose }: { anchor: MiniBarAnchor; onClose: () => void }) {
  const ctx = useBelge();
  const { editor, fmt, styles } = ctx;
  const [hl, setHl] = useState<string | null>("#FFFF00");
  const [color, setColor] = useState<string | null>("#C00000");
  const chain = () => editor.chain().focus();
  const run = (f: () => void) => () => {
    f();
    if (anchor.context) onClose();
  };
  return (
    <>
      <MiniToolbarRow>
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
      </MiniToolbarRow>
      <MiniToolbarRow>
        <RibbonButton size="small" icon={<TextBold20Regular />} tip="Kalın" shortcut="Ctrl+B" checked={fmt.bold} onClick={run(() => chain().toggleBold().run())} />
        <RibbonButton size="small" icon={<TextItalic20Regular />} tip="İtalik" shortcut="Ctrl+I" checked={fmt.italic} onClick={run(() => chain().toggleItalic().run())} />
        <RibbonButton size="small" icon={<TextUnderline20Regular />} tip="Altı Çizili" shortcut="Ctrl+U" checked={!!fmt.underline} onClick={run(() => chain().toggleUnderline().run())} />
        <ColorSplitButton icon={<Highlight20Regular />} tip="Metin Vurgu Rengi" color={hl} variant="highlight" noneLabel="Renk Yok" onColorChange={setHl} onApply={(c) => (c ? chain().setHighlight({ color: c }).run() : chain().unsetHighlight().run())} />
        <ColorSplitButton icon={<TextColor20Regular />} tip="Yazı Tipi Rengi" color={color} noneLabel="Otomatik" onColorChange={setColor} onApply={(c) => (c ? chain().setColor(c).run() : chain().unsetColor().run())} />
        <RibbonButton size="small" icon={<TextBulletListLtr20Regular />} tip="Madde İşaretleri" checked={fmt.list === "bullet"} onClick={run(() => applyBulletMain(editor))} />
        <RibbonButton size="small" icon={<TextNumberListLtr20Regular />} tip="Numaralandırma" checked={fmt.list === "ordered"} onClick={run(() => applyNumberMain(editor))} />
        {ctx.inBody && <RibbonButton size="small" icon={<CommentAdd20Regular />} tip="Yeni Açıklama" onClick={run(ctx.newComment)} />}
      </MiniToolbarRow>
    </>
  );
}

/** Sağlayıcı içinde çizilen mini çubuk: belge düzenleyicisinin fare seçimini izler; bağlam menüsü `controlRef` ile açar. */
export function MiniToolbarHost({ controlRef }: { controlRef: { current: ((a: MiniBarAnchor | null) => void) | null } }) {
  const ed = useBelge().editor;
  return (
    <UiMiniToolbarHost
      className="bl-minibar"
      controlRef={controlRef}
      root={() => (ed.isDestroyed ? null : (ed.view.dom as HTMLElement))}
      selectionRect={() => {
        if (ed.isDestroyed) return null;
        // Çift/üçlü tıklamada seçim DOM'da değişmiş ama düzenleyiciye henüz okunmamış olabilir.
        if (ed.view.hasFocus()) syncDomSelection(ed.view);
        const sel = ed.state.selection;
        if (sel.empty || !ed.view.hasFocus() || !ed.state.doc.textBetween(sel.from, sel.to).trim()) return null;
        return window.getSelection()?.getRangeAt(0)?.getBoundingClientRect();
      }}
    >
      {({ anchor, close }) => <BelgeMiniBar anchor={anchor} onClose={close} />}
    </UiMiniToolbarHost>
  );
}
