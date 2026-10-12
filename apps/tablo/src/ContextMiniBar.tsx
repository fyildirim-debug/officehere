/**
 * Hücre (ve satır/sütun başlığı) bağlam menüsünün üstündeki mini araç çubuğu. Menünün parçası olarak
 * çizilir: menü kapanınca o da kapanır; çubukla çalışılırken menü gizlenir. Komutlar Giriş sekmesinin
 * biçim eylemleridir.
 */
import { useEffect, useRef, useState } from "react";
import {
  BorderAll20Regular,
  BorderBottom20Regular,
  BorderLeft20Regular,
  BorderNone20Regular,
  BorderOutside20Regular,
  BorderOutsideThick20Regular,
  BorderRight20Regular,
  BorderTop20Regular,
  DecimalArrowLeft20Regular,
  DecimalArrowRight20Regular,
  FontDecrease20Regular,
  FontIncrease20Regular,
  Money20Regular,
  PaintBucket20Regular,
  TableCellsMerge20Regular,
  TextAlignCenter20Regular,
  TextAlignLeft20Regular,
  TextAlignRight20Regular,
  TextBold20Regular,
  TextColor20Regular,
  TextItalic20Regular,
} from "@fluentui/react-icons";
import { FONT_LIST, FONT_SIZES, fontStack } from "@officehere/core";
import { ColorSplitButton, Menu, MenuItem, MiniToolbar, MiniToolbarRow, RibbonButton, RibbonCombo, RibbonSplitButton, useMenuClose, type MiniBarAnchor } from "@officehere/ui";
import type { Controller } from "./controller";
import * as F from "./actions/format";
import { DEFAULT_FONT, DEFAULT_SIZE } from "./model/styles";
import { useThemePalette } from "./themeContext";

const BORDERS: { kind: string; label: string; icon: React.ReactElement }[] = [
  { kind: "bottom", label: "Alt Kenarlık", icon: <BorderBottom20Regular /> },
  { kind: "top", label: "Üst Kenarlık", icon: <BorderTop20Regular /> },
  { kind: "left", label: "Sol Kenarlık", icon: <BorderLeft20Regular /> },
  { kind: "right", label: "Sağ Kenarlık", icon: <BorderRight20Regular /> },
  { kind: "none", label: "Kenarlık Yok", icon: <BorderNone20Regular /> },
  { kind: "all", label: "Tüm Kenarlıklar", icon: <BorderAll20Regular /> },
  { kind: "outside", label: "Dış Kenarlıklar", icon: <BorderOutside20Regular /> },
  { kind: "thickOutside", label: "Kalın Dış Kenarlıklar", icon: <BorderOutsideThick20Regular /> },
];

/** Metin simgesi (yüzde, binlik ayırıcı). */
function Glyph({ text, size }: { text: string; size: number }) {
  return (
    <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true">
      <text x="10" y="14.5" textAnchor="middle" fontSize={size} fontWeight="600" fill="currentColor" fontFamily="Segoe UI, sans-serif">
        {text}
      </text>
    </svg>
  );
}

export function ContextMiniBar({ ctl }: { ctl: Controller }) {
  const closeMenu = useMenuClose();
  const palette = useThemePalette();
  const mark = useRef<HTMLSpanElement>(null);
  const [anchor, setAnchor] = useState<MiniBarAnchor | null>(null);
  const [fill, setFill] = useState<string | null>("#FFFF00");
  const [color, setColor] = useState<string | null>("#C00000");
  const [border, setBorder] = useState(BORDERS[0]);
  const menuEl = () => mark.current?.closest<HTMLElement>("[data-oh-popover]") ?? null;

  // Menü yerleşince (taşmaya göre yukarı/sola açılmış olabilir) çubuk onun üstüne konur.
  useEffect(() => {
    const id = requestAnimationFrame(() => {
      const r = menuEl()?.getBoundingClientRect();
      if (r && r.left > -1000) setAnchor({ x: r.left, y: r.top, context: true });
    });
    return () => cancelAnimationFrame(id);
  }, []);

  const st = F.activeStyle(ctl);
  const { row, col } = ctl.sel.active;
  const merged = !!ctl.sheet.mergeAt(row, col);
  /** Komut uygulanır, menü ve çubuk kapanır, odak ızgaraya döner. */
  const run = (f: () => void) => () => {
    f();
    closeMenu();
    requestAnimationFrame(() => ctl.focusGrid());
  };
  const hideMenu = () => menuEl()?.setAttribute("data-oh-minibar-hide", "");

  return (
    <>
      <span ref={mark} hidden />
      <MiniToolbar anchor={anchor} onClose={closeMenu} layer={2} fade={false} className="tb-minibar">
        <div onPointerDownCapture={hideMenu} style={{ display: "contents" }}>
          <MiniToolbarRow>
            <RibbonCombo value={st.font ?? DEFAULT_FONT} width={116} options={FONT_LIST.map((f) => ({ value: f, style: { fontFamily: fontStack(f) } }))} onCommit={(f) => run(() => F.setFont(ctl, f))()} tip="Yazı Tipi" />
            <RibbonCombo value={String(st.size ?? DEFAULT_SIZE).replace(".", ",")} width={46} options={FONT_SIZES.map((s) => String(s).replace(".", ","))} onCommit={(s) => run(() => F.setFontSize(ctl, Number(s.replace(",", "."))))()} tip="Yazı Tipi Boyutu" />
            <RibbonButton size="small" icon={<FontIncrease20Regular />} tip="Yazı Tipi Boyutunu Büyüt" onClick={run(() => F.growFont(ctl, 1))} />
            <RibbonButton size="small" icon={<FontDecrease20Regular />} tip="Yazı Tipi Boyutunu Küçült" onClick={run(() => F.growFont(ctl, -1))} />
            <RibbonButton size="small" icon={<Money20Regular />} tip="Finansal Sayı Biçimi" onClick={run(() => F.accounting(ctl))} />
            <RibbonButton size="small" icon={<Glyph text="%" size={13} />} tip="Yüzde Stili" onClick={run(() => F.percent(ctl))} />
            <RibbonButton size="small" icon={<Glyph text="000" size={10} />} tip="Binlik Ayırıcı Stili" onClick={run(() => F.comma(ctl))} />
          </MiniToolbarRow>
          <MiniToolbarRow>
            <RibbonButton size="small" icon={<TextBold20Regular />} tip="Kalın" shortcut="Ctrl+B" checked={!!st.bold} onClick={run(() => F.toggle(ctl, "bold"))} />
            <RibbonButton size="small" icon={<TextItalic20Regular />} tip="İtalik" shortcut="Ctrl+I" checked={!!st.italic} onClick={run(() => F.toggle(ctl, "italic"))} />
            <RibbonButton size="small" icon={<TextAlignLeft20Regular />} tip="Sola Hizala" checked={st.hAlign === "left"} onClick={run(() => F.setHAlign(ctl, "left"))} />
            <RibbonButton size="small" icon={<TextAlignCenter20Regular />} tip="Ortala" checked={st.hAlign === "center"} onClick={run(() => F.setHAlign(ctl, "center"))} />
            <RibbonButton size="small" icon={<TextAlignRight20Regular />} tip="Sağa Hizala" checked={st.hAlign === "right"} onClick={run(() => F.setHAlign(ctl, "right"))} />
            <RibbonSplitButton
              size="small"
              icon={border.icon}
              tip={border.label}
              onClick={run(() => F.borders(ctl, border.kind))}
              menu={
                <Menu width={200}>
                  {BORDERS.map((b) => (
                    <MenuItem key={b.kind} icon={b.icon} label={b.label} onClick={run(() => (setBorder(b), F.borders(ctl, b.kind)))} />
                  ))}
                </Menu>
              }
            />
            <ColorSplitButton icon={<PaintBucket20Regular />} tip="Dolgu Rengi" color={fill} theme={palette} noneLabel="Dolgu Yok" onColorChange={setFill} onApply={(c) => run(() => F.setFill(ctl, c))()} />
            <ColorSplitButton icon={<TextColor20Regular />} tip="Yazı Tipi Rengi" color={color} theme={palette} noneLabel="Otomatik" onColorChange={setColor} onApply={(c) => run(() => F.setColor(ctl, c))()} />
            <RibbonButton size="small" icon={<DecimalArrowRight20Regular />} tip="Ondalık Artır" onClick={run(() => F.changeDecimals(ctl, 1))} />
            <RibbonButton size="small" icon={<DecimalArrowLeft20Regular />} tip="Ondalık Azalt" onClick={run(() => F.changeDecimals(ctl, -1))} />
            <RibbonButton size="small" icon={<TableCellsMerge20Regular />} tip="Birleştir ve Ortala" checked={merged} onClick={run(() => F.mergeCells(ctl, "center"))} />
          </MiniToolbarRow>
        </div>
      </MiniToolbar>
    </>
  );
}
