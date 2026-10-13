/** Giriş sekmesi: Pano, Yazı Tipi, Hizalama, Sayı grupları. */
import {
  AlignBottom20Regular,
  AlignCenterVertical20Regular,
  AlignTop20Regular,
  BorderAll20Regular,
  BorderBottom20Regular,
  BorderBottomDouble20Regular,
  BorderBottomThick20Regular,
  BorderInside20Regular,
  BorderLeft20Regular,
  BorderNone20Regular,
  BorderOutside20Regular,
  BorderOutsideThick20Regular,
  BorderRight20Regular,
  BorderTop20Regular,
  ClipboardPaste32Regular,
  Copy20Regular,
  Cut20Regular,
  DecimalArrowLeft20Regular,
  DecimalArrowRight20Regular,
  FontDecrease20Regular,
  FontIncrease20Regular,
  Money20Regular,
  PaintBrush20Regular,
  PaintBucket20Regular,
  TableCellsMerge20Regular,
  TextAlignCenter20Regular,
  TextAlignLeft20Regular,
  TextAlignRight20Regular,
  TextBold20Regular,
  TextColor20Regular,
  TextDirectionRotate270Right20Regular,
  TextDirectionRotate315Right20Regular,
  TextDirectionRotate45Right20Regular,
  TextIndentDecrease20Regular,
  TextIndentIncrease20Regular,
  TextItalic20Regular,
  TextUnderline20Regular,
  TextWrap20Regular,
  TextStrikethrough20Regular,
  ClipboardPaste20Regular,
  TextDirectionHorizontalLtr20Regular,
  TextFont24Regular,
  TextAlignLeft24Regular,
  NumberSymbol24Regular,
} from "@fluentui/react-icons";
import { useState } from "react";
import { FONT_LIST, FONT_SIZES, fontStack } from "@officehere/core";
import { ColorPalette, ColorSplitButton, Menu, MenuHeader, MenuItem, MenuSeparator, RibbonButton, RibbonCombo, RibbonGroup, RibbonMenuButton, RibbonRow, RibbonSplitButton, RibbonStack } from "@officehere/ui";
import { formatValue } from "@officehere/formula";
import type { Controller } from "../controller";
import * as F from "../actions/format";
import * as E from "../actions/edit";
import type { PasteOptions } from "../model/clipboard";
import { DEFAULT_FONT, DEFAULT_SIZE, NUMBER_FORMATS } from "../model/styles";
import { StylesGroup, CellsGroup, EditingGroup } from "./HomeTab2";
import { useThemePalette } from "../themeContext";

export function PasteMenu({ ctl }: { ctl: Controller }) {
  const p = (what: PasteOptions["what"], transpose = false) => () => E.pasteFromButton(ctl, { what, transpose });
  return (
    <Menu width={230}>
      <MenuHeader>Yapıştır</MenuHeader>
      <MenuItem label="Yapıştır" shortcut="Ctrl+V" onClick={() => E.pasteFromButton(ctl)} />
      <MenuItem label="Formüller" onClick={p("formulas")} />
      <MenuItem label="Formüller ve Sayı Biçimlendirme" onClick={p("formulasAndFormats")} />
      <MenuItem label="Kaynak Biçimlendirmesini Koru" onClick={p("all")} />
      <MenuItem label="Kenarlık Yok" onClick={p("allExceptBorders")} />
      <MenuItem label="Kaynak Sütun Genişliklerini Koru" onClick={() => (E.pasteFromButton(ctl, { what: "all" }), E.pasteFromButton(ctl, { what: "widths" }))} />
      <MenuItem label="Devrik Dönüşüm" onClick={p("all", true)} />
      <MenuHeader>Değerleri Yapıştır</MenuHeader>
      <MenuItem label="Değerler" onClick={p("values")} />
      <MenuItem label="Değerler ve Sayı Biçimlendirmesi" onClick={p("valuesAndFormats")} />
      <MenuHeader>Diğer Yapıştırma Seçenekleri</MenuHeader>
      <MenuItem label="Biçimlendirme" onClick={p("formats")} />
      <MenuItem label="Bağlantıyı Yapıştır" description="Kopyalanan hücrelere başvuran formüller yapıştırır; kaynak değişince güncellenir." onClick={p("link")} />
      <MenuSeparator />
      <MenuItem label="Özel Yapıştır..." shortcut="Ctrl+Alt+V" onClick={() => ctl.openDialog("pasteSpecial")} />
    </Menu>
  );
}

const BORDERS: { kind: string; label: string; icon: React.ReactElement }[] = [
  { kind: "bottom", label: "Alt Kenarlık", icon: <BorderBottom20Regular /> },
  { kind: "top", label: "Üst Kenarlık", icon: <BorderTop20Regular /> },
  { kind: "left", label: "Sol Kenarlık", icon: <BorderLeft20Regular /> },
  { kind: "right", label: "Sağ Kenarlık", icon: <BorderRight20Regular /> },
  { kind: "none", label: "Kenarlık Yok", icon: <BorderNone20Regular /> },
  { kind: "all", label: "Tüm Kenarlıklar", icon: <BorderAll20Regular /> },
  { kind: "outside", label: "Dış Kenarlıklar", icon: <BorderOutside20Regular /> },
  { kind: "thickOutside", label: "Kalın Dış Kenarlıklar", icon: <BorderOutsideThick20Regular /> },
  { kind: "doubleBottom", label: "Çift Alt Kenarlık", icon: <BorderBottomDouble20Regular /> },
  { kind: "thickBottom", label: "Kalın Alt Kenarlık", icon: <BorderBottomThick20Regular /> },
  { kind: "topBottom", label: "Üst ve Alt Kenarlık", icon: <BorderInside20Regular /> },
  { kind: "topThickBottom", label: "Üst ve Kalın Alt Kenarlık", icon: <BorderBottomThick20Regular /> },
  { kind: "topDoubleBottom", label: "Üst ve Çift Alt Kenarlık", icon: <BorderBottomDouble20Regular /> },
];

export function HomeTab({ ctl }: { ctl: Controller }) {
  const st = F.activeStyle(ctl);
  const palette = useThemePalette();
  const [fontColor, setFontColor] = useState<string | null>("#C00000");
  const [fillColor, setFillColor] = useState<string | null>("#FFFF00");
  const [lastBorder, setLastBorder] = useState(BORDERS[0]);
  const [borderColor, setBorderColor] = useState("#000000");
  const [borderStyle, setBorderStyle] = useState<"thin" | "medium" | "thick" | "dashed" | "dotted" | "double">("thin");
  const { row, col } = ctl.sel.active;
  const v = ctl.wb.getValue(ctl.sheet, row, col);
  const fmtLabel = (code: string) => {
    if (code === "General") return "Genel";
    const sample = typeof v === "number" ? v : null;
    return sample === null ? "" : formatValue(sample, code).text.trim();
  };
  const curFmt = NUMBER_FORMATS.find((f) => f.code === (st.numFmt ?? "General"));
  const merged = !!ctl.sheet.mergeAt(row, col);

  return (
    <>
      <RibbonGroup label="Pano" icon={<ClipboardPaste20Regular />}>
        <RibbonSplitButton size="large" icon={<ClipboardPaste32Regular />} label="Yapıştır" shortcut="Ctrl+V" description="Panodaki içeriği ekler." onClick={() => E.pasteFromButton(ctl)} menu={<PasteMenu ctl={ctl} />} />
        <RibbonStack>
          <RibbonButton icon={<Cut20Regular />} label="Kes" shortcut="Ctrl+X" onClick={() => copyBtn(ctl, true)} />
          <RibbonButton icon={<Copy20Regular />} label="Kopyala" shortcut="Ctrl+C" onClick={() => copyBtn(ctl, false)} />
          <RibbonButton icon={<PaintBrush20Regular />} label="Biçim Boyacısı" checked={!!ctl.painter} description="Seçimdeki biçimlendirmeyi kopyalayıp başka hücrelere uygular." onClick={() => (ctl.painter ? ctl.cancelModes() : ctl.startPainter(false))} />
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Yazı Tipi" icon={<TextFont24Regular />} onLauncher={() => ctl.openDialog("formatCells", { tab: "font" })} launcherTip="Hücreleri Biçimlendir: Yazı Tipi">
        <RibbonStack>
          <RibbonRow>
            <RibbonCombo value={st.font ?? DEFAULT_FONT} width={124} options={FONT_LIST.map((f) => ({ value: f, style: { fontFamily: fontStack(f) } }))} onCommit={(f) => (F.setFont(ctl, f), requestAnimationFrame(() => ctl.focusGrid()))} tip="Yazı Tipi" shortcut="Ctrl+Shift+F" />
            <RibbonCombo value={String(st.size ?? DEFAULT_SIZE).replace(".", ",")} width={46} options={FONT_SIZES.map((s) => String(s).replace(".", ","))} onCommit={(s) => (F.setFontSize(ctl, Number(s.replace(",", "."))), requestAnimationFrame(() => ctl.focusGrid()))} tip="Yazı Tipi Boyutu" />
            <RibbonButton size="small" icon={<FontIncrease20Regular />} tip="Yazı Tipi Boyutunu Büyüt" shortcut="Ctrl+Shift+>" onClick={() => F.growFont(ctl, 1)} />
            <RibbonButton size="small" icon={<FontDecrease20Regular />} tip="Yazı Tipi Boyutunu Küçült" shortcut="Ctrl+Shift+<" onClick={() => F.growFont(ctl, -1)} />
          </RibbonRow>
          <RibbonRow>
            <RibbonButton size="small" icon={<TextBold20Regular />} tip="Kalın" shortcut="Ctrl+B" checked={!!st.bold} onClick={() => F.toggle(ctl, "bold")} />
            <RibbonButton size="small" icon={<TextItalic20Regular />} tip="İtalik" shortcut="Ctrl+I" checked={!!st.italic} onClick={() => F.toggle(ctl, "italic")} />
            <RibbonSplitButton
              size="small"
              icon={<TextUnderline20Regular />}
              tip="Altı Çizili"
              shortcut="Ctrl+U"
              checked={!!st.underline}
              onClick={() => F.toggleUnderline(ctl)}
              menu={
                <Menu>
                  <MenuItem icon={<TextUnderline20Regular />} label="Altı Çizili" checked={st.underline === "single"} onClick={() => F.toggleUnderline(ctl, "single")} />
                  <MenuItem label="Çift Alt Çizgi" checked={st.underline === "double"} onClick={() => F.toggleUnderline(ctl, "double")} />
                </Menu>
              }
            />
            <RibbonButton size="small" icon={<TextStrikethrough20Regular />} tip="Üstü Çizili" shortcut="Ctrl+5" checked={!!st.strike} onClick={() => F.toggle(ctl, "strike")} />
            <span className="oh-ribbon-sep-small" />
            <RibbonSplitButton
              size="small"
              icon={lastBorder.icon}
              tip={lastBorder.label}
              onClick={() => F.borders(ctl, lastBorder.kind, borderColor, borderStyle)}
              menu={(close) => (
                <Menu width={220}>
                  <MenuHeader>Kenarlıklar</MenuHeader>
                  {BORDERS.map((b) => (
                    <MenuItem
                      key={b.kind}
                      icon={b.icon}
                      label={b.label}
                      onClick={() => {
                        setLastBorder(b);
                        F.borders(ctl, b.kind, borderColor, borderStyle);
                        close();
                      }}
                    />
                  ))}
                  <MenuHeader>Kenarlık Çiz</MenuHeader>
                  <MenuItem
                    label="Çizgi Rengi"
                    submenu={
                      <div className="oh-menu oh-menu-palette">
                        <ColorPalette value={borderColor} theme={palette} onSelect={(c) => setBorderColor(c ?? "#000000")} noneLabel="Otomatik" />
                      </div>
                    }
                  />
                  <MenuItem
                    label="Çizgi Stili"
                    submenu={
                      <Menu>
                        {(
                          [
                            ["thin", "İnce"],
                            ["medium", "Orta"],
                            ["thick", "Kalın"],
                            ["dashed", "Kesik"],
                            ["dotted", "Noktalı"],
                            ["double", "Çift"],
                          ] as const
                        ).map(([k, l]) => (
                          <MenuItem key={k} label={l} checked={borderStyle === k} onClick={() => setBorderStyle(k)} />
                        ))}
                      </Menu>
                    }
                  />
                  <MenuSeparator />
                  <MenuItem label="Diğer Kenarlıklar..." onClick={() => ctl.openDialog("formatCells", { tab: "border" })} />
                </Menu>
              )}
            />
            <ColorSplitButton icon={<PaintBucket20Regular />} tip="Dolgu Rengi" description="Seçili hücrelerin arka plan rengini değiştirir." color={fillColor} theme={palette} noneLabel="Dolgu Yok" onColorChange={setFillColor} onApply={(c) => F.setFill(ctl, c)} />
            <ColorSplitButton icon={<TextColor20Regular />} tip="Yazı Tipi Rengi" color={fontColor} theme={palette} noneLabel="Otomatik" onColorChange={setFontColor} onApply={(c) => F.setColor(ctl, c)} />
          </RibbonRow>
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Hizalama" icon={<TextAlignLeft24Regular />} onLauncher={() => ctl.openDialog("formatCells", { tab: "align" })} launcherTip="Hücreleri Biçimlendir: Hizalama">
        <RibbonStack>
          <RibbonRow>
            <RibbonButton size="small" icon={<AlignTop20Regular />} tip="Üste Hizala" checked={st.vAlign === "top"} onClick={() => F.setVAlign(ctl, "top")} />
            <RibbonButton size="small" icon={<AlignCenterVertical20Regular />} tip="Ortaya Hizala" checked={st.vAlign === "middle"} onClick={() => F.setVAlign(ctl, "middle")} />
            <RibbonButton size="small" icon={<AlignBottom20Regular />} tip="Alta Hizala" checked={!st.vAlign || st.vAlign === "bottom"} onClick={() => F.setVAlign(ctl, "bottom")} />
            <RibbonMenuButton
              size="small"
              icon={<TextDirectionRotate315Right20Regular />}
              tip="Yönlendirme"
              menu={
                <Menu width={220}>
                  <MenuItem icon={<TextDirectionRotate315Right20Regular />} label="Saatin Tersi Yönünde Açı" checked={st.rotation === 45} onClick={() => F.setRotation(ctl, 45)} />
                  <MenuItem icon={<TextDirectionRotate45Right20Regular />} label="Saat Yönünde Açı" checked={st.rotation === 135} onClick={() => F.setRotation(ctl, 135)} />
                  <MenuItem icon={<TextDirectionHorizontalLtr20Regular />} label="Dikey Metin" checked={st.rotation === 255} onClick={() => F.setRotation(ctl, 255)} />
                  <MenuItem icon={<TextDirectionRotate270Right20Regular />} label="Metni Yukarı Döndür" checked={st.rotation === 90} onClick={() => F.setRotation(ctl, 90)} />
                  <MenuItem label="Metni Aşağı Döndür" checked={st.rotation === 180} onClick={() => F.setRotation(ctl, 180)} />
                  <MenuSeparator />
                  <MenuItem label="Hücre Hizalamasını Biçimlendir" onClick={() => ctl.openDialog("formatCells", { tab: "align" })} />
                </Menu>
              }
            />
          </RibbonRow>
          <RibbonRow>
            <RibbonButton size="small" icon={<TextAlignLeft20Regular />} tip="Sola Hizala" checked={st.hAlign === "left"} onClick={() => F.setHAlign(ctl, "left")} />
            <RibbonButton size="small" icon={<TextAlignCenter20Regular />} tip="Ortala" checked={st.hAlign === "center"} onClick={() => F.setHAlign(ctl, "center")} />
            <RibbonButton size="small" icon={<TextAlignRight20Regular />} tip="Sağa Hizala" checked={st.hAlign === "right"} onClick={() => F.setHAlign(ctl, "right")} />
            <RibbonButton size="small" icon={<TextIndentDecrease20Regular />} tip="Girintiyi Azalt" onClick={() => F.indent(ctl, -1)} />
            <RibbonButton size="small" icon={<TextIndentIncrease20Regular />} tip="Girintiyi Artır" onClick={() => F.indent(ctl, 1)} />
          </RibbonRow>
        </RibbonStack>
        <RibbonStack>
          <RibbonButton icon={<TextWrap20Regular />} label="Metni Kaydır" checked={!!st.wrap} description="Çok uzun metni birden çok satıra kaydırarak görünür yapar." onClick={() => F.toggleWrap(ctl)} />
          <RibbonSplitButton
            icon={<TableCellsMerge20Regular />}
            label="Birleştir ve Ortala"
            checked={merged}
            onClick={() => F.mergeCells(ctl, "center")}
            menu={
              <Menu>
                <MenuItem icon={<TableCellsMerge20Regular />} label="Birleştir ve Ortala" onClick={() => F.mergeCells(ctl, "center")} />
                <MenuItem label="Arasında Birleştir" onClick={() => F.mergeCells(ctl, "across")} />
                <MenuItem label="Hücreleri Birleştir" onClick={() => F.mergeCells(ctl, "merge")} />
                <MenuItem label="Hücrelerin Birleştirmesini Kaldır" onClick={() => F.mergeCells(ctl, "unmerge")} />
              </Menu>
            }
          />
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Sayı" icon={<NumberSymbol24Regular />} onLauncher={() => ctl.openDialog("formatCells", { tab: "number" })} launcherTip="Hücreleri Biçimlendir: Sayı">
        <RibbonStack>
          <RibbonCombo
            value={curFmt?.label ?? "Özel"}
            width={150}
            editable={false}
            listHeight={420}
            options={NUMBER_FORMATS.map((f) => ({ value: f.label, label: `${f.label}${fmtLabel(f.code) && f.id !== "general" && f.id !== "text" ? `   ${fmtLabel(f.code)}` : ""}` }))}
            onCommit={(label) => {
              const f = NUMBER_FORMATS.find((x) => x.label === label);
              if (f) F.setNumFmt(ctl, f.code);
            }}
            tip="Sayı Biçimi"
          />
          <RibbonRow>
            <RibbonSplitButton
              size="small"
              icon={<Money20Regular />}
              tip="Finansal Sayı Biçimi"
              description="Seçili hücreler için ₺ gibi alternatif bir para birimi biçimi seçin."
              onClick={() => F.accounting(ctl)}
              menu={
                <Menu>
                  <MenuItem label="₺ Türkçe (Türkiye)" onClick={() => F.accounting(ctl)} />
                  <MenuItem label="$ İngilizce (ABD)" onClick={() => F.setNumFmt(ctl, '_-[$$-409]* #,##0.00_ ;_-[$$-409]* -#,##0.00 ;_-[$$-409]* "-"??_ ;_-@_ ')} />
                  <MenuItem label="€ Euro (€ 123)" onClick={() => F.setNumFmt(ctl, '_-* #,##0.00 [$€-x-euro1]_-;-* #,##0.00 [$€-x-euro1]_-;_-* "-"?? [$€-x-euro1]_-;_-@_-')} />
                  <MenuItem label="£ İngilizce (Birleşik Krallık)" onClick={() => F.setNumFmt(ctl, '_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-')} />
                  <MenuSeparator />
                  <MenuItem label="Diğer Finansal Biçimler..." onClick={() => ctl.openDialog("formatCells", { tab: "number" })} />
                </Menu>
              }
            />
            <RibbonButton size="small" icon={<PercentGlyph />} tip="Yüzde Stili" shortcut="Ctrl+Shift+%" onClick={() => F.percent(ctl)} />
            <RibbonButton size="small" icon={<ThousandGlyph />} tip="Binlik Ayırıcı Stili" onClick={() => F.comma(ctl)} />
            <RibbonButton size="small" icon={<DecimalArrowRight20Regular />} tip="Ondalık Artır" onClick={() => F.changeDecimals(ctl, 1)} />
            <RibbonButton size="small" icon={<DecimalArrowLeft20Regular />} tip="Ondalık Azalt" onClick={() => F.changeDecimals(ctl, -1)} />
          </RibbonRow>
        </RibbonStack>
      </RibbonGroup>

      <StylesGroup ctl={ctl} />
      <CellsGroup ctl={ctl} />
      <EditingGroup ctl={ctl} />
    </>
  );
}

function copyBtn(ctl: Controller, cut: boolean) {
  const d = E.copySelection(ctl, cut);
  if (d) void E.writeSystemClipboard(d);
}

function PercentGlyph() {
  return (
    <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true">
      <text x="10" y="15" textAnchor="middle" fontSize="13" fontWeight="600" fill="currentColor" fontFamily="Segoe UI, sans-serif">
        %
      </text>
    </svg>
  );
}

function ThousandGlyph() {
  return (
    <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true">
      <text x="10" y="14" textAnchor="middle" fontSize="10" fontWeight="600" fill="currentColor" fontFamily="Segoe UI, sans-serif">
        000
      </text>
    </svg>
  );
}
