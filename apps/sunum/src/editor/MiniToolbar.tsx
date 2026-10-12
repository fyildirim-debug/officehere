/**
 * Mini araç çubuğu: metin kutusunda fareyle metin seçilince seçimin üstünde, sağ tıkta bağlam
 * menüsünün üstünde açılır. Komutlar şeritteki Giriş sekmesinin eylemlerini çağırır.
 */
import { useEffect, useState, type RefObject } from "react";
import {
  FontDecrease20Regular,
  FontIncrease20Regular,
  Highlight20Regular,
  TextAlignCenter20Regular,
  TextAlignLeft20Regular,
  TextAlignRight20Regular,
  TextBold20Regular,
  TextBulletListLtr20Regular,
  TextColor20Regular,
  TextItalic20Regular,
  TextUnderline20Regular,
} from "@fluentui/react-icons";
import { FONT_LIST, fontStack } from "@officehere/core";
import { MiniToolbar, MiniToolbarRow, RibbonButton, RibbonCombo, useMiniToolbar, type MiniBarAnchor } from "@officehere/ui";
import { resolveFont } from "../model/textStyle";
import type { Color } from "../model/types";
import { ThemeColorSplit } from "../ribbon/common";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";
import { applyRun, common, growFont, setListStyle, setParaProps, SLIDE_FONT_SIZES, textInfo } from "../state/textActions";
import { isInEditedShape } from "./contextKind";

type Flag = "bold" | "italic" | "underline";

function SunumMiniBar({ state, actions: a, anchor, onClose }: { state: AppState; actions: Actions; anchor: MiniBarAnchor; onClose: () => void }) {
  const [hl, setHl] = useState<Color | null>({ rgb: "#FFFF00" });
  const [fontColor, setFontColor] = useState<Color | null>({ rgb: "#FF0000" });
  const theme = state.pres.theme;
  const info = textInfo(a);
  const runs = info?.runs ?? [];
  const para = info?.para;
  const fontName = common<string>(runs, "font");
  const size = common<number>(runs, "size");
  const all = (k: Flag) => runs.length > 0 && runs.every((r) => !!r[k]);
  // Bağlam kipinde komut uygulanınca çubuk kapanır (menüyle birlikte); seçim kipinde açık kalır.
  const run = (f: () => void) => () => {
    f();
    if (anchor.context) onClose();
  };
  const toggle = (k: Flag) => run(() => {
    const on = !all(k);
    applyRun(a, (p) => ({ ...p, [k]: on }));
  });
  const fontOptions = [
    { value: theme.fonts.major, label: `${theme.fonts.major} (Başlıklar)`, group: "Tema Yazı Tipleri", style: { fontFamily: fontStack(theme.fonts.major) } },
    { value: theme.fonts.minor, label: `${theme.fonts.minor} (Gövde)`, group: "Tema Yazı Tipleri", style: { fontFamily: fontStack(theme.fonts.minor) } },
    ...FONT_LIST.map((f) => ({ value: f, group: "Tüm Yazı Tipleri", style: { fontFamily: fontStack(f) } })),
  ];
  const bulleted = para?.bullet?.type === "char";
  return (
    <>
      <MiniToolbarRow>
        <RibbonCombo
          value={fontName ? resolveFont(fontName, theme) : ""}
          width={124}
          options={fontOptions}
          onCommit={(v) => applyRun(a, (p) => ({ ...p, font: v === theme.fonts.major ? "+mj" : v === theme.fonts.minor ? "+mn" : v }))}
          tip="Yazı Tipi"
        />
        <RibbonCombo
          value={size !== undefined ? String(Math.round(size * 10) / 10).replace(".", ",") : ""}
          width={50}
          options={SLIDE_FONT_SIZES.map((v) => String(v).replace(".", ","))}
          onCommit={(v) => {
            const n = parseFloat(v.replace(",", "."));
            if (n > 0 && n < 4000) applyRun(a, (p) => ({ ...p, size: n }));
          }}
          tip="Yazı Tipi Boyutu"
        />
        <RibbonButton size="small" icon={<FontIncrease20Regular />} tip="Yazı Tipi Boyutunu Büyüt" shortcut="Ctrl+]" onClick={() => growFont(a, 1)} />
        <RibbonButton size="small" icon={<FontDecrease20Regular />} tip="Yazı Tipi Boyutunu Küçült" shortcut="Ctrl+[" onClick={() => growFont(a, -1)} />
        <RibbonButton size="small" icon={<TextBulletListLtr20Regular />} tip="Madde İşaretleri" checked={bulleted} onClick={run(() => setListStyle(a, bulleted ? { type: "none" } : { type: "char", char: "•", font: "Arial" }, !para?.marL))} />
      </MiniToolbarRow>
      <MiniToolbarRow>
        <RibbonButton size="small" icon={<TextBold20Regular />} tip="Kalın" shortcut="Ctrl+B" checked={all("bold")} onClick={toggle("bold")} />
        <RibbonButton size="small" icon={<TextItalic20Regular />} tip="İtalik" shortcut="Ctrl+I" checked={all("italic")} onClick={toggle("italic")} />
        <RibbonButton size="small" icon={<TextUnderline20Regular />} tip="Altı Çizili" shortcut="Ctrl+U" checked={all("underline")} onClick={toggle("underline")} />
        <ThemeColorSplit icon={<Highlight20Regular />} tip="Metin Vurgu Rengi" theme={theme} color={hl} noneLabel="Vurgu Yok" onColorChange={setHl} onApply={(c) => applyRun(a, (p) => ({ ...p, highlight: c ?? undefined }))} />
        <ThemeColorSplit icon={<TextColor20Regular />} tip="Yazı Tipi Rengi" theme={theme} color={fontColor} autoLabel="Otomatik" onColorChange={setFontColor} onApply={(c) => applyRun(a, (p) => ({ ...p, color: c ?? undefined }))} />
        <RibbonButton size="small" icon={<TextAlignLeft20Regular />} tip="Sola Hizala" shortcut="Ctrl+L" checked={(para?.align ?? "l") === "l"} onClick={run(() => setParaProps(a, { align: "l" }))} />
        <RibbonButton size="small" icon={<TextAlignCenter20Regular />} tip="Ortala" shortcut="Ctrl+E" checked={para?.align === "ctr"} onClick={run(() => setParaProps(a, { align: "ctr" }))} />
        <RibbonButton size="small" icon={<TextAlignRight20Regular />} tip="Sağa Hizala" shortcut="Ctrl+R" checked={para?.align === "r"} onClick={run(() => setParaProps(a, { align: "r" }))} />
      </MiniToolbarRow>
    </>
  );
}

/** Slayt alanında bir kez çizilir; yalnız metin düzenlenirken açılır. */
export function SunumMiniToolbar({ state, actions, slideRef }: { state: AppState; actions: Actions; slideRef: RefObject<HTMLDivElement | null> }) {
  const editor = () => slideRef.current?.querySelector<HTMLElement>(".tb-edit") ?? null;
  const { anchor, close } = useMiniToolbar({
    root: editor,
    selectionRect: () => {
      const el = editor();
      const sel = window.getSelection();
      if (!el || document.activeElement !== el || !sel || sel.isCollapsed || !sel.rangeCount || !el.contains(sel.anchorNode)) return null;
      return sel.toString().trim() ? sel.getRangeAt(0).getBoundingClientRect() : null;
    },
    // Düzenlenen metinde sağ tık (Shift ile tarayıcı menüsü açılır, çubuk açılmaz).
    contextTarget: (t) => !!actions.s.editing && (!!t.closest(".tb-edit") || isInEditedShape(t, actions.s, slideRef.current)),
  });
  const editing = !!state.editing;
  useEffect(() => {
    if (!editing) close();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editing]);
  return (
    <MiniToolbar anchor={editing ? anchor : null} onClose={close} className="sn-minibar">
      {anchor && <SunumMiniBar state={state} actions={actions} anchor={anchor} onClose={close} />}
    </MiniToolbar>
  );
}
