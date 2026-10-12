/** Giriş sekmesi: Pano, Slaytlar, Yazı Tipi, Paragraf, Çizim, Düzenleme. */
import { useState } from "react";
import {
  ArrowReset20Regular,
  ArrowSwap20Regular,
  ClipboardPaste32Regular,
  Copy20Regular,
  Cut20Regular,
  FontDecrease20Regular,
  FontIncrease20Regular,
  FontSpaceTrackingOut20Regular,
  Highlight20Regular,
  LayerDiagonal20Regular,
  PaintBrush20Regular,
  Search20Regular,
  SelectAllOn20Regular,
  SlideAdd32Regular,
  SlideLayout20Regular,
  TextAlignCenter20Regular,
  TextAlignJustify20Regular,
  TextAlignLeft20Regular,
  TextAlignRight20Regular,
  TextBold20Regular,
  TextBulletListLtr20Regular,
  TextChangeCase20Regular,
  TextClearFormatting20Regular,
  TextColor20Regular,
  TextDirectionRotate90Right20Regular,
  TextIndentDecreaseLtr20Regular,
  TextIndentIncreaseLtr20Regular,
  TextItalic20Regular,
  TextLineSpacing20Regular,
  TextNumberListLtr20Regular,
  TextStrikethrough20Regular,
  TextUnderline20Regular,
  TextboxAlignMiddle20Regular,
  Shapes32Regular,
  CursorClick20Regular,
  TextT20Regular,
  ClipboardPaste20Regular,
  Copy24Regular,
  ClipboardPaste24Regular,
  Search24Regular,
  Shapes24Regular,
  SlideAdd24Regular,
  TextFont24Regular,
  TextParagraph24Regular,
} from "@fluentui/react-icons";
import { FONT_LIST, FONT_SIZES, fontStack } from "@officehere/core";
import { Menu, MenuHeader, MenuItem, MenuSeparator, RibbonButton, RibbonCombo, RibbonGroup, RibbonMenuButton, RibbonRow, RibbonSplitButton, RibbonStack, useMenuClose } from "@officehere/ui";
import { resolveFont } from "../model/textStyle";
import type { BulletDef, Color } from "../model/types";
import { captureFormat } from "../state/painter";
import { applyCase, applyRun, changeLevel, common, growFont, setBodyProps, setListStyle, setParaProps } from "../state/textActions";
import { autoNumber } from "../render/textHtml";
import { InlineShapeGallery, LayoutGallery, ThemeColorSplit } from "./common";
import { BULLET_PRESETS, NUMBER_PRESETS, SPACING_PRESETS } from "./presets";
import { ArrangeMenu, QuickStylesMenu, ShapeEffectsMenu, ShapeFillButton, ShapeOutlineButton } from "./shapeMenus";
import type { RibbonCtx } from "./types";

function BulletGrid({ onPick, kind }: { onPick: (b: BulletDef) => void; kind: "bullet" | "number" }) {
  const close = useMenuClose();
  const tiles =
    kind === "bullet"
      ? BULLET_PRESETS.map((b) => ({ label: b.label, def: b.bullet, sample: b.bullet.type === "char" ? b.bullet.char : "Yok" }))
      : NUMBER_PRESETS.map((n) => ({ label: n.label, def: (n.scheme ? { type: "auto", scheme: n.scheme } : { type: "none" }) as BulletDef, sample: n.scheme ? [1, 2, 3].map((i) => autoNumber(n.scheme!, i)).join("\n") : "Yok" }));
  return (
    <div className="oh-menu">
      <div className="preset-grid" style={{ gridTemplateColumns: "repeat(4, 64px)" }}>
        {tiles.map((t) => (
          <button
            key={t.label}
            type="button"
            className="preset-btn"
            title={t.label}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              onPick(t.def);
              close();
            }}
            style={{ width: 64, height: 64, border: "1px solid var(--oh-border)", flexDirection: "column", fontSize: t.sample === "Yok" ? 12 : 15, whiteSpace: "pre", lineHeight: 1.15 }}
          >
            {kind === "bullet" && t.sample !== "Yok" ? `${t.sample} ——\n${t.sample} ——\n${t.sample} ——` : t.sample.replace(/\n/g, " ——\n") + (t.sample !== "Yok" ? " ——" : "")}
          </button>
        ))}
      </div>
    </div>
  );
}

export function HomeTab({ ctx }: { ctx: RibbonCtx }) {
  const { state, actions: a, text, cmd } = ctx;
  const theme = state.pres.theme;
  const [fontColor, setFontColor] = useState<Color | null>({ rgb: "#FF0000" });
  const [hl, setHl] = useState<Color | null>({ rgb: "#FFFF00" });
  const runs = text?.runs ?? [];
  const fontName = common<string>(runs, "font");
  const size = common<number>(runs, "size");
  const has = !!text;
  const para = text?.para;
  const bu = para?.bullet;
  const hasSel = state.sel.length > 0;
  const slide = state.pres.slides[state.slideIdx];
  const all = (k: "bold" | "italic" | "underline" | "strike" | "shadow") => runs.length > 0 && runs.every((r) => !!r[k]);
  const toggle = (k: "bold" | "italic" | "underline" | "strike" | "shadow") => {
    const on = !all(k);
    applyRun(a, (p) => ({ ...p, [k]: on }));
  };
  const fontOptions = [
    { value: theme.fonts.major, label: `${theme.fonts.major} (Başlıklar)`, group: "Tema Yazı Tipleri", style: { fontFamily: fontStack(theme.fonts.major) } },
    { value: theme.fonts.minor, label: `${theme.fonts.minor} (Gövde)`, group: "Tema Yazı Tipleri", style: { fontFamily: fontStack(theme.fonts.minor) } },
    ...FONT_LIST.map((f) => ({ value: f, group: "Tüm Yazı Tipleri", style: { fontFamily: fontStack(f) } })),
  ];

  return (
    <>
      <RibbonGroup label="Pano" icon={<ClipboardPaste24Regular />} onLauncher={() => cmd.paste()} launcherTip="Pano">
        <RibbonSplitButton
          size="large"
          icon={<ClipboardPaste32Regular />}
          label="Yapıştır"
          shortcut="Ctrl+V"
          tip="Yapıştır"
          description="Panodaki içeriği ekleyin."
          onClick={() => cmd.paste()}
          menu={
            <Menu>
              <MenuHeader>Yapıştırma Seçenekleri:</MenuHeader>
              <MenuItem icon={<ClipboardPaste20Regular />} label="Hedef Temayı Kullan" onClick={() => cmd.paste("dest")} />
              <MenuItem icon={<PaintBrush20Regular />} label="Kaynak Biçimlendirmesini Koru" onClick={() => cmd.paste("keep")} />
              <MenuItem icon={<TextT20Regular />} label="Yalnızca Metni Koru" onClick={() => cmd.paste("text")} />
            </Menu>
          }
        />
        <RibbonStack>
          <RibbonButton icon={<Cut20Regular />} label="Kes" shortcut="Ctrl+X" onClick={cmd.cut} disabled={!hasSel && !state.editing && state.focus !== "slides"} />
          <RibbonSplitButton
            icon={<Copy20Regular />}
            label="Kopyala"
            shortcut="Ctrl+C"
            onClick={cmd.copy}
            menu={
              <Menu>
                <MenuItem icon={<Copy24Regular />} label="Kopyala" shortcut="Ctrl+C" onClick={cmd.copy} />
                <MenuItem label="Çoğalt" shortcut="Ctrl+D" onClick={() => (hasSel ? a.duplicateSelection() : a.duplicateSlides())} />
              </Menu>
            }
          />
          <RibbonButton
            icon={<PaintBrush20Regular />}
            label="Biçim Boyacısı"
            checked={!!state.painter}
            tip="Biçim Boyacısı"
            description="Bir nesnenin biçimini kopyalayıp başka bir nesneye uygulayın. Kalıcı kullanım için Ctrl+Shift+C / Ctrl+Shift+V."
            disabled={!hasSel}
            onClick={() => (state.painter ? a.store.set({ painter: null }) : captureFormat(a, false))}
          />
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Slaytlar" icon={<SlideAdd24Regular />}>
        <RibbonSplitButton
          size="large"
          icon={<SlideAdd32Regular />}
          label="Yeni Slayt"
          shortcut="Ctrl+M"
          description="Sunuya slayt ekleyin."
          onClick={() => a.newSlide()}
          menu={
            <>
              <LayoutGallery pres={state.pres} header="Tema" onPick={(id) => a.newSlide(id)} />
              <div className="oh-menu" style={{ marginTop: -4, borderTop: 0, borderRadius: "0 0 6px 6px" }}>
                <MenuItem label="Seçili Slaytları Çoğalt" onClick={() => a.duplicateSlides()} />
              </div>
            </>
          }
        />
        <RibbonStack>
          <RibbonMenuButton icon={<SlideLayout20Regular />} label="Düzen" tip="Düzen" description="Seçili slaydın düzenini değiştirin." disabled={!slide} menu={<LayoutGallery pres={state.pres} selectedId={slide?.layoutId} onPick={(id) => a.setLayout(id)} />} />
          <RibbonButton icon={<ArrowReset20Regular />} label="Sıfırla" tip="Sıfırla" description="Slayt yer tutucularının konumunu, boyutunu ve biçimini varsayılana sıfırlayın." disabled={!slide} onClick={() => a.resetSlides()} />
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Yazı Tipi" icon={<TextFont24Regular />} onLauncher={() => ctx.openDialog("font")} launcherTip="Yazı Tipi">
        <RibbonStack>
          <RibbonRow>
            <RibbonCombo
              value={fontName ? resolveFont(fontName, theme) : ""}
              width={140}
              options={fontOptions}
              onCommit={(v) => applyRun(a, (p) => ({ ...p, font: v === theme.fonts.major ? "+mj" : v === theme.fonts.minor ? "+mn" : v }))}
              tip="Yazı Tipi"
              shortcut="Ctrl+Shift+F"
              disabled={!has}
            />
            <RibbonCombo
              value={size !== undefined ? String(Math.round(size * 10) / 10).replace(".", ",") : ""}
              width={52}
              options={FONT_SIZES.map((v) => String(v).replace(".", ","))}
              onCommit={(v) => {
                const n = parseFloat(v.replace(",", "."));
                if (n > 0 && n < 4000) applyRun(a, (p) => ({ ...p, size: n }));
              }}
              tip="Yazı Tipi Boyutu"
              shortcut="Ctrl+Shift+P"
              disabled={!has}
            />
            <RibbonButton size="small" icon={<FontIncrease20Regular />} tip="Yazı Tipi Boyutunu Büyüt" shortcut="Ctrl+]" disabled={!has} onClick={() => growFont(a, 1)} />
            <RibbonButton size="small" icon={<FontDecrease20Regular />} tip="Yazı Tipi Boyutunu Küçült" shortcut="Ctrl+[" disabled={!has} onClick={() => growFont(a, -1)} />
            <RibbonButton size="small" icon={<TextClearFormatting20Regular />} tip="Tüm Biçimlendirmeyi Temizle" description="Seçimdeki tüm biçimlendirmeyi kaldırıp yalnızca normal metin bırakın." disabled={!has} onClick={() => applyRun(a, (p) => (p.field ? { field: p.field } : {}))} />
          </RibbonRow>
          <RibbonRow>
            <RibbonButton size="small" icon={<TextBold20Regular />} tip="Kalın" shortcut="Ctrl+B" checked={all("bold")} disabled={!has} onClick={() => toggle("bold")} />
            <RibbonButton size="small" icon={<TextItalic20Regular />} tip="İtalik" shortcut="Ctrl+I" checked={all("italic")} disabled={!has} onClick={() => toggle("italic")} />
            <RibbonButton size="small" icon={<TextUnderline20Regular />} tip="Altı Çizili" shortcut="Ctrl+U" checked={all("underline")} disabled={!has} onClick={() => toggle("underline")} />
            <RibbonButton
              size="small"
              icon={<span style={{ fontWeight: 700, fontSize: 14, textShadow: "1px 1px 1px rgba(0,0,0,.45)", width: 16, display: "inline-block", textAlign: "center" }}>S</span>}
              tip="Metin Gölgesi"
              description="Metnin arkasına gölge ekleyin."
              checked={all("shadow")}
              disabled={!has}
              onClick={() => toggle("shadow")}
            />
            <RibbonButton size="small" icon={<TextStrikethrough20Regular />} tip="Üstü Çizili" checked={all("strike")} disabled={!has} onClick={() => toggle("strike")} />
            <RibbonMenuButton
              size="small"
              icon={<FontSpaceTrackingOut20Regular />}
              tip="Karakter Aralığı"
              disabled={!has}
              menu={
                <Menu>
                  {SPACING_PRESETS.map((s) => (
                    <MenuItem key={s.label} label={s.label} checked={common<number>(runs, "spacing") === s.value} onClick={() => applyRun(a, (p) => ({ ...p, spacing: s.value }))} />
                  ))}
                  <MenuSeparator />
                  <MenuItem label="Diğer Aralıklar..." onClick={() => ctx.openDialog("font")} />
                </Menu>
              }
            />
            <RibbonMenuButton
              size="small"
              icon={<TextChangeCase20Regular />}
              tip="Büyük/Küçük Harf Değiştir"
              disabled={!has}
              menu={
                <Menu>
                  <MenuItem label="Tümcedeki ilk harf büyük." onClick={() => applyCase(a, "sentence")} />
                  <MenuItem label="küçük harf" onClick={() => applyCase(a, "lower")} />
                  <MenuItem label="BÜYÜK HARF" onClick={() => applyCase(a, "upper")} />
                  <MenuItem label="Her Sözcüğün İlk Harfi Büyük" onClick={() => applyCase(a, "title")} />
                  <MenuItem label="bÜYÜK/kÜÇÜK hARF dEĞİŞTİR" onClick={() => applyCase(a, "toggle")} />
                </Menu>
              }
            />
            <ThemeColorSplit icon={<Highlight20Regular />} tip="Metin Vurgu Rengi" theme={theme} color={hl} noneLabel="Vurgu Yok" disabled={!has} onColorChange={setHl} onApply={(c) => applyRun(a, (p) => ({ ...p, highlight: c ?? undefined }))} />
            <ThemeColorSplit icon={<TextColor20Regular />} tip="Yazı Tipi Rengi" theme={theme} color={fontColor} autoLabel="Otomatik" disabled={!has} onColorChange={setFontColor} onApply={(c) => applyRun(a, (p) => ({ ...p, color: c ?? undefined }))} />
          </RibbonRow>
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Paragraf" icon={<TextParagraph24Regular />} onLauncher={() => ctx.openDialog("paragraph")} launcherTip="Paragraf">
        <RibbonStack>
          <RibbonRow>
            <RibbonSplitButton
              icon={<TextBulletListLtr20Regular />}
              tip="Madde İşaretleri"
              checked={bu?.type === "char"}
              disabled={!has}
              onClick={() => setListStyle(a, bu?.type === "char" ? { type: "none" } : { type: "char", char: "•", font: "Arial" }, !para?.marL)}
              menu={<BulletGrid kind="bullet" onPick={(b) => setListStyle(a, b, !para?.marL)} />}
            />
            <RibbonSplitButton
              icon={<TextNumberListLtr20Regular />}
              tip="Numaralandırma"
              checked={bu?.type === "auto"}
              disabled={!has}
              onClick={() => setListStyle(a, bu?.type === "auto" ? { type: "none" } : { type: "auto", scheme: "arabicPeriod" })}
              menu={<BulletGrid kind="number" onPick={(b) => setListStyle(a, b)} />}
            />
            <RibbonButton size="small" icon={<TextIndentDecreaseLtr20Regular />} tip="Liste Düzeyini Azalt" shortcut="Shift+Tab" disabled={!has} onClick={() => changeLevel(a, -1)} />
            <RibbonButton size="small" icon={<TextIndentIncreaseLtr20Regular />} tip="Liste Düzeyini Artır" shortcut="Tab" disabled={!has} onClick={() => changeLevel(a, 1)} />
            <RibbonMenuButton
              size="small"
              icon={<TextLineSpacing20Regular />}
              tip="Satır Aralığı"
              disabled={!has}
              menu={
                <Menu>
                  {[1, 1.5, 2, 2.5, 3].map((v) => (
                    <MenuItem key={v} label={v.toLocaleString("tr-TR", { minimumFractionDigits: 1 })} checked={Math.abs((para?.lineSpacing ?? 100) / 100 - v) < 0.01 || (v === 1 && (para?.lineSpacing ?? 100) === 90)} onClick={() => setParaProps(a, { lineSpacing: v * 100, lineSpacingPt: undefined })} />
                  ))}
                  <MenuSeparator />
                  <MenuItem label="Satır Aralığı Seçenekleri..." onClick={() => ctx.openDialog("paragraph")} />
                </Menu>
              }
            />
          </RibbonRow>
          <RibbonRow>
            <RibbonButton size="small" icon={<TextAlignLeft20Regular />} tip="Sola Hizala" shortcut="Ctrl+L" checked={has && (para?.align ?? "l") === "l"} disabled={!has} onClick={() => setParaProps(a, { align: "l" })} />
            <RibbonButton size="small" icon={<TextAlignCenter20Regular />} tip="Ortala" shortcut="Ctrl+E" checked={para?.align === "ctr"} disabled={!has} onClick={() => setParaProps(a, { align: "ctr" })} />
            <RibbonButton size="small" icon={<TextAlignRight20Regular />} tip="Sağa Hizala" shortcut="Ctrl+R" checked={para?.align === "r"} disabled={!has} onClick={() => setParaProps(a, { align: "r" })} />
            <RibbonButton size="small" icon={<TextAlignJustify20Regular />} tip="İki Yana Yasla" shortcut="Ctrl+J" checked={para?.align === "just"} disabled={!has} onClick={() => setParaProps(a, { align: "just" })} />
            <RibbonMenuButton
              size="small"
              icon={<TextDirectionRotate90Right20Regular />}
              tip="Metin Yönü"
              disabled={!hasSel}
              menu={
                <Menu>
                  <MenuItem label="Yatay" checked={(text?.body?.vert ?? "horz") === "horz"} onClick={() => setBodyProps(a, { vert: "horz" })} />
                  <MenuItem label="Tüm metni 90° döndür" checked={text?.body?.vert === "vert"} onClick={() => setBodyProps(a, { vert: "vert" })} />
                  <MenuItem label="Tüm metni 270° döndür" checked={text?.body?.vert === "vert270"} onClick={() => setBodyProps(a, { vert: "vert270" })} />
                  <MenuItem label="Yığılmış" checked={text?.body?.vert === "stacked"} onClick={() => setBodyProps(a, { vert: "stacked" })} />
                </Menu>
              }
            />
            <RibbonMenuButton
              size="small"
              icon={<TextboxAlignMiddle20Regular />}
              tip="Metni Hizala"
              description="Metnin metin kutusu içindeki dikey hizalamasını değiştirin."
              disabled={!hasSel}
              menu={
                <Menu>
                  <MenuItem label="Üst" checked={(text?.body?.anchor ?? "t") === "t"} onClick={() => setBodyProps(a, { anchor: "t" })} />
                  <MenuItem label="Orta" checked={text?.body?.anchor === "ctr"} onClick={() => setBodyProps(a, { anchor: "ctr" })} />
                  <MenuItem label="Alt" checked={text?.body?.anchor === "b"} onClick={() => setBodyProps(a, { anchor: "b" })} />
                </Menu>
              }
            />
          </RibbonRow>
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Çizim" icon={<Shapes24Regular />} onLauncher={() => cmd.togglePane("shape")} launcherTip="Şekli Biçimlendir">
        <InlineShapeGallery onPick={cmd.pickShape} />
        <RibbonMenuButton size="large" icon={<Shapes32Regular />} label="Yerleştir" tip="Yerleştir" description="Nesneleri sıralayın, gruplayın, hizalayın ve döndürün." menu={(close) => <ArrangeMenu a={a} onSelectionPane={() => (close(), cmd.togglePane("selection"))} onMoreRotate={() => (close(), ctx.openDialog("sizePos"))} />} />
        <QuickStylesMenu a={a} theme={theme} />
        <RibbonStack>
          <ShapeFillButton a={a} theme={theme} label="Şekil Dolgusu" size="medium" />
          <ShapeOutlineButton a={a} theme={theme} label="Şekil Anahattı" size="medium" />
          <ShapeEffectsMenu a={a} theme={theme} label="Şekil Efektleri" size="medium" />
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Düzenleme" icon={<Search24Regular />}>
        <RibbonStack>
          <RibbonButton icon={<Search20Regular />} label="Bul" shortcut="Ctrl+F" onClick={() => ctx.openDialog("find")} />
          <RibbonSplitButton
            icon={<ArrowSwap20Regular />}
            label="Değiştir"
            shortcut="Ctrl+H"
            onClick={() => ctx.openDialog("replace")}
            menu={
              <Menu>
                <MenuItem label="Değiştir..." shortcut="Ctrl+H" onClick={() => ctx.openDialog("replace")} />
                <MenuItem label="Yazı Tiplerini Değiştir..." onClick={() => ctx.openDialog("replaceFonts")} />
              </Menu>
            }
          />
          <RibbonMenuButton
            icon={<CursorClick20Regular />}
            label="Seç"
            menu={
              <Menu>
                <MenuItem icon={<SelectAllOn20Regular />} label="Tümünü Seç" shortcut="Ctrl+A" onClick={() => a.selectAll()} />
                <MenuItem icon={<LayerDiagonal20Regular />} label="Seçim Bölmesi..." onClick={() => cmd.togglePane("selection")} />
              </Menu>
            }
          />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}

