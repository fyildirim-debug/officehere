// Giriş sekmesi: Pano, Yazı Tipi, Paragraf, Stiller, Düzenleme.
import { useState } from "react";
import {
  ArrowSort20Regular,
  ArrowSwap20Regular,
  BorderAll20Regular,
  BorderBottom20Regular,
  BorderLeft20Regular,
  BorderNone20Regular,
  BorderOutside20Regular,
  BorderRight20Regular,
  BorderTop20Regular,
  ClipboardPaste32Regular,
  Copy20Regular,
  Cut20Regular,
  Cursor20Regular,
  Highlight20Regular,
  LineHorizontal120Regular,
  PaintBrush20Regular,
  PaintBucket20Regular,
  Search20Regular,
  TextAlignCenter20Regular,
  TextAlignJustify20Regular,
  TextAlignLeft20Regular,
  TextAlignRight20Regular,
  TextBold20Regular,
  TextBulletListLtr20Regular,
  TextBulletListTree20Regular,
  TextChangeCase20Regular,
  TextClearFormatting20Regular,
  TextColor20Regular,
  TextFontSize20Regular,
  TextIndentDecreaseLtr20Regular,
  TextIndentIncreaseLtr20Regular,
  TextItalic20Regular,
  TextNumberListLtr20Regular,
  TextParagraph20Regular,
  TextStrikethrough20Regular,
  TextSubscript20Regular,
  TextSuperscript20Regular,
  TextUnderline20Regular,
  TextLineSpacing20Regular,
  ArrowUp20Regular,
  ArrowDown20Regular,
  TextEditStyle20Regular,
} from "@fluentui/react-icons";
import { FONT_LIST, FONT_SIZES, fontStack } from "@officehere/core";
import {
  ColorPalette,
  ColorSplitButton,
  Menu,
  MenuItem,
  MenuSeparator,
  RibbonButton,
  RibbonCombo,
  RibbonGallery,
  RibbonGroup,
  RibbonMenuButton,
  RibbonRow,
  RibbonSplitButton,
  RibbonStack,
} from "@officehere/ui";
import { useBelge } from "../state/context";
import { applyStyle, changeCase, changeIndent, clearFormatting, setFontSize, setLineSpacing, setParaAttrs, stepFontSize, currentPara } from "../editor/commands";
import { UNDERLINE_TYPES, type UnderlineType } from "../editor/marks";
import { execClipboard, pasteFromClipboard } from "./clipboard";
import { BulletLibrary, MultilevelLibrary, NumberLibrary, applyBulletMain, applyNumberMain } from "./lists";
import { previewStyle } from "../model/styles";
import type { ParaBorders } from "../editor/paragraphFormat";

const GALLERY_STYLES = [
  "Normal",
  "NoSpacing",
  "Heading1",
  "Heading2",
  "Heading3",
  "Title",
  "Subtitle",
  "SubtleEmphasis",
  "Emphasis",
  "IntenseEmphasis",
  "Strong",
  "Quote",
  "IntenseQuote",
  "SubtleReference",
  "IntenseReference",
  "BookTitle",
  "ListParagraph",
  "Heading4",
  "Heading5",
  "Heading6",
];

const LINE_OPTIONS = [1, 1.15, 1.5, 2, 2.5, 3];

export function HomeTab() {
  const ctx = useBelge();
  const { editor, fmt, styles } = ctx;
  const [fontColor, setFontColor] = useState<string | null>("#C00000");
  const [hl, setHl] = useState<string | null>("#FFFF00");
  const [shade, setShade] = useState<string | null>(null);

  const chain = () => editor.chain().focus();

  const applyShading = (c: string | null) => {
    if (fmt.inTable && (fmt.cellSelection || fmt.empty)) {
      chain().setCellAttribute("background", c).run();
      return;
    }
    const p = currentPara(editor);
    const sel = editor.state.selection;
    const whole = sel.empty || (p && sel.from <= p.pos + 1 && sel.to >= p.pos + p.node.nodeSize - 1) || sel.$from.parent !== sel.$to.parent;
    if (whole) setParaAttrs(editor, { shading: c });
    else if (c) chain().setBackgroundColor(c).run();
    else chain().unsetBackgroundColor().run();
  };

  const border = (kind: "bottom" | "top" | "left" | "right" | "none" | "all" | "outside" | "hr") => {
    if (kind === "hr") {
      chain().setHorizontalRule().run();
      return;
    }
    if (fmt.inTable) {
      chain()
        .updateAttributes("table", { tableStyle: kind === "none" ? "none" : "grid" })
        .run();
      return;
    }
    const b = { style: "single" as const, size: 0.5, color: "#000000", space: 1 };
    const p = currentPara(editor);
    const cur = ((p?.node.attrs.borders as ParaBorders | null) ?? {}) as ParaBorders;
    let next: ParaBorders | null;
    switch (kind) {
      case "none":
        next = null;
        break;
      case "all":
      case "outside":
        next = { top: b, bottom: b, left: b, right: b };
        break;
      default:
        next = { ...cur, [kind]: cur[kind] ? undefined : b };
        if (!Object.values(next).some(Boolean)) next = null;
    }
    setParaAttrs(editor, { borders: next });
  };

  const underlineMenu = (
    <Menu>
      {UNDERLINE_TYPES.map((u) => (
        <MenuItem
          key={u.id}
          label={<span className="bl-u-sample" style={{ textDecorationStyle: u.id === "wave" ? "wavy" : u.id === "thick" ? "solid" : (u.id as "solid"), textDecorationThickness: u.id === "thick" ? 2 : undefined }}>{u.label}</span>}
          checked={fmt.underline === u.id}
          onClick={() => chain().setUnderlineType({ type: u.id as UnderlineType }).run()}
        />
      ))}
      <MenuSeparator />
      <MenuItem
        label="Alt Çizgi Rengi"
        submenu={
          <div className="oh-menu oh-menu-palette">
            <ColorPalette noneLabel="Otomatik" onSelect={(c) => chain().setUnderlineType({ type: (fmt.underline as UnderlineType) ?? "single", color: c }).run()} />
          </div>
        }
      />
    </Menu>
  );

  const pasteMenu = (
    <Menu>
      <MenuItem label="Kaynak Biçimlendirmesini Koru (K)" onClick={() => pasteFromClipboard(editor, "keep", ctx.notify)} />
      <MenuItem label="Biçimlendirmeyi Birleştir (B)" onClick={() => pasteFromClipboard(editor, "merge", ctx.notify)} />
      <MenuItem label="Yalnızca Metni Koru (M)" onClick={() => pasteFromClipboard(editor, "text", ctx.notify)} />
    </Menu>
  );

  const caseMenu = (
    <Menu>
      <MenuItem label="Cümle düzeni." onClick={() => changeCase(editor, "sentence")} />
      <MenuItem label="küçük harf" onClick={() => changeCase(editor, "lower")} />
      <MenuItem label="BÜYÜK HARF" onClick={() => changeCase(editor, "upper")} />
      <MenuItem label="Her Sözcüğün İlk Harfi Büyük" onClick={() => changeCase(editor, "title")} />
      <MenuItem label="bÜYÜK/kÜÇÜK hARF dEĞİŞTİR" onClick={() => changeCase(editor, "toggle")} />
    </Menu>
  );

  const before = fmt.spaceBefore ?? 0;
  const after = fmt.spaceAfter ?? (fmt.styleId === "NoSpacing" ? 0 : 8);
  const lineMenu = (
    <Menu>
      {LINE_OPTIONS.map((v) => (
        <MenuItem key={v} label={v.toLocaleString("tr-TR", { minimumFractionDigits: 1 })} checked={(fmt.lineRule ?? "auto") === "auto" && Math.abs((fmt.lineHeight ?? 1.08) - v) < 0.01} onClick={() => setLineSpacing(editor, v)} />
      ))}
      <MenuSeparator />
      <MenuItem label="Satır Aralığı Seçenekleri..." onClick={() => ctx.openDialog("paragraph")} />
      <MenuSeparator />
      {before > 0 ? (
        <MenuItem icon={<ArrowUp20Regular />} label="Paragraftan Önce Boşluğu Kaldır" onClick={() => setParaAttrs(editor, { spaceBefore: 0 })} />
      ) : (
        <MenuItem icon={<ArrowUp20Regular />} label="Paragraftan Önce Boşluk Ekle" onClick={() => setParaAttrs(editor, { spaceBefore: 12 })} />
      )}
      {after > 0 ? (
        <MenuItem icon={<ArrowDown20Regular />} label="Paragraftan Sonra Boşluğu Kaldır" onClick={() => setParaAttrs(editor, { spaceAfter: 0 })} />
      ) : (
        <MenuItem icon={<ArrowDown20Regular />} label="Paragraftan Sonra Boşluk Ekle" onClick={() => setParaAttrs(editor, { spaceAfter: 12 })} />
      )}
    </Menu>
  );

  const borderMenu = (
    <Menu>
      <MenuItem icon={<BorderBottom20Regular />} label="Alt Kenarlık" onClick={() => border("bottom")} />
      <MenuItem icon={<BorderTop20Regular />} label="Üst Kenarlık" onClick={() => border("top")} />
      <MenuItem icon={<BorderLeft20Regular />} label="Sol Kenarlık" onClick={() => border("left")} />
      <MenuItem icon={<BorderRight20Regular />} label="Sağ Kenarlık" onClick={() => border("right")} />
      <MenuSeparator />
      <MenuItem icon={<BorderNone20Regular />} label="Kenarlık Yok" onClick={() => border("none")} />
      <MenuItem icon={<BorderAll20Regular />} label="Tüm Kenarlıklar" onClick={() => border("all")} />
      <MenuItem icon={<BorderOutside20Regular />} label="Dış Kenarlıklar" onClick={() => border("outside")} />
      <MenuSeparator />
      <MenuItem icon={<LineHorizontal120Regular />} label="Yatay Çizgi" onClick={() => border("hr")} />
    </Menu>
  );

  const alignBtn = (a: string, icon: React.ReactElement, tip: string, key: string) => (
    <RibbonButton size="small" icon={icon} tip={tip} shortcut={key} checked={fmt.align === a} onClick={() => chain().setTextAlign(a).run()} />
  );

  return (
    <>
      <RibbonGroup label="Pano" icon={<ClipboardPaste32Regular />}>
        <RibbonSplitButton size="large" icon={<ClipboardPaste32Regular />} label="Yapıştır" shortcut="Ctrl+V" onClick={() => pasteFromClipboard(editor, "keep", ctx.notify)} menu={pasteMenu} />
        <RibbonStack>
          <RibbonButton icon={<Cut20Regular />} label="Kes" shortcut="Ctrl+X" disabled={fmt.empty} onClick={() => execClipboard(editor, "cut")} />
          <RibbonButton icon={<Copy20Regular />} label="Kopyala" shortcut="Ctrl+C" disabled={fmt.empty} onClick={() => execClipboard(editor, "copy")} />
          <span onDoubleClick={() => ctx.startFormatPainter(true)}>
            <RibbonButton
              icon={<PaintBrush20Regular />}
              label="Biçim Boyacısı"
              shortcut="Ctrl+Shift+C"
              checked={ctx.formatPainter.active}
              description="Biçimlendirmeyi bir yerden kopyalayıp başka bir yere uygulayın. Birden fazla yere uygulamak için çift tıklayın."
              onClick={() => ctx.startFormatPainter(false)}
            />
          </span>
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Yazı Tipi" onLauncher={() => ctx.openDialog("font")} launcherTip="Yazı Tipi Ayarları (Ctrl+D)" icon={<TextFontSize20Regular />}>
        <RibbonStack>
          <RibbonRow>
            <RibbonCombo
              value={fmt.font}
              width={140}
              options={FONT_LIST.map((f) => ({ value: f, style: { fontFamily: fontStack(f), fontSize: 14 } }))}
              onCommit={(v) => chain().setFontFamily(v).run()}
              tip="Yazı Tipi"
              shortcut="Ctrl+Shift+F"
            />
            <RibbonCombo
              value={String(fmt.size).replace(".", ",")}
              width={52}
              options={FONT_SIZES.map((s) => String(s).replace(".", ","))}
              onCommit={(v) => {
                const n = parseFloat(v.replace(",", "."));
                if (Number.isFinite(n) && n > 0) setFontSize(editor, n);
              }}
              tip="Yazı Tipi Boyutu"
              shortcut="Ctrl+Shift+P"
            />
            <RibbonButton size="small" icon={<span className="bl-grow">A<sup>▲</sup></span>} tip="Yazı Tipini Büyült" shortcut="Ctrl+Shift+>" onClick={() => stepFontSize(editor, styles, 1)} />
            <RibbonButton size="small" icon={<span className="bl-grow is-shrink">A<sup>▼</sup></span>} tip="Yazı Tipini Küçült" shortcut="Ctrl+Shift+<" onClick={() => stepFontSize(editor, styles, -1)} />
            <RibbonMenuButton size="small" icon={<TextChangeCase20Regular />} tip="Büyük/Küçük Harf Değiştir" shortcut="Shift+F3" menu={caseMenu} />
            <RibbonButton size="small" icon={<TextClearFormatting20Regular />} tip="Tüm Biçimlendirmeyi Temizle" description="Seçimdeki tüm biçimlendirmeyi kaldırarak yalnızca normal, biçimlendirilmemiş metni bırakın." onClick={() => clearFormatting(editor)} />
          </RibbonRow>
          <RibbonRow>
            <RibbonButton size="small" icon={<TextBold20Regular />} tip="Kalın" shortcut="Ctrl+B" checked={fmt.bold} onClick={() => chain().toggleBold().run()} />
            <RibbonButton size="small" icon={<TextItalic20Regular />} tip="İtalik" shortcut="Ctrl+I" checked={fmt.italic} onClick={() => chain().toggleItalic().run()} />
            <RibbonSplitButton size="small" icon={<TextUnderline20Regular />} tip="Altı Çizili" shortcut="Ctrl+U" checked={!!fmt.underline} onClick={() => chain().toggleUnderline().run()} menu={underlineMenu} />
            <RibbonButton size="small" icon={<TextStrikethrough20Regular />} tip="Üstü Çizili" checked={fmt.strike} onClick={() => chain().toggleStrike().run()} />
            <RibbonButton size="small" icon={<TextSubscript20Regular />} tip="Alt Simge" shortcut="Ctrl+=" checked={fmt.sub} onClick={() => chain().toggleSubscript().run()} />
            <RibbonButton size="small" icon={<TextSuperscript20Regular />} tip="Üst Simge" shortcut="Ctrl+Shift++" checked={fmt.sup} onClick={() => chain().toggleSuperscript().run()} />
            <ColorSplitButton
              icon={<Highlight20Regular />}
              tip="Metin Vurgu Rengi"
              color={hl}
              variant="highlight"
              noneLabel="Renk Yok"
              onColorChange={setHl}
              onApply={(c) => (c ? chain().setHighlight({ color: c }).run() : chain().unsetHighlight().run())}
            />
            <ColorSplitButton icon={<TextColor20Regular />} tip="Yazı Tipi Rengi" color={fontColor} noneLabel="Otomatik" onColorChange={setFontColor} onApply={(c) => (c ? chain().setColor(c).run() : chain().unsetColor().run())} />
          </RibbonRow>
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Paragraf" onLauncher={() => ctx.openDialog("paragraph")} launcherTip="Paragraf Ayarları" icon={<TextParagraph20Regular />}>
        <RibbonStack>
          <RibbonRow>
            <RibbonSplitButton
              size="small"
              icon={<TextBulletListLtr20Regular />}
              tip="Madde İşaretleri"
              checked={fmt.list === "bullet"}
              onClick={() => applyBulletMain(editor)}
              menu={<BulletLibrary editor={editor} current={fmt.list === "bullet" ? fmt.listStyle : null} />}
            />
            <RibbonSplitButton
              size="small"
              icon={<TextNumberListLtr20Regular />}
              tip="Numaralandırma"
              checked={fmt.list === "ordered"}
              onClick={() => applyNumberMain(editor)}
              menu={<NumberLibrary editor={editor} current={fmt.list === "ordered" ? fmt.listStyle : null} />}
            />
            <RibbonMenuButton size="small" icon={<TextBulletListTree20Regular />} tip="Çok Düzeyli Liste" menu={<MultilevelLibrary editor={editor} />} />
            <RibbonButton size="small" icon={<TextIndentDecreaseLtr20Regular />} tip="Girintiyi Azalt" shortcut="Ctrl+Shift+M" onClick={() => changeIndent(editor, -1)} />
            <RibbonButton size="small" icon={<TextIndentIncreaseLtr20Regular />} tip="Girintiyi Artır" shortcut="Ctrl+M" onClick={() => changeIndent(editor, 1)} />
            <RibbonButton size="small" icon={<ArrowSort20Regular />} tip="Sırala" description="Geçerli seçimi alfabetik ya da sayısal olarak sıralayın." onClick={() => ctx.openDialog("sort")} />
            <RibbonButton size="small" icon={<TextParagraph20Regular />} tip="Tümünü Göster" shortcut="Ctrl+*" checked={ctx.view.showMarks} onClick={() => ctx.setView({ showMarks: !ctx.view.showMarks })} />
          </RibbonRow>
          <RibbonRow>
            {alignBtn("left", <TextAlignLeft20Regular />, "Sola Hizala", "Ctrl+L")}
            {alignBtn("center", <TextAlignCenter20Regular />, "Ortala", "Ctrl+E")}
            {alignBtn("right", <TextAlignRight20Regular />, "Sağa Hizala", "Ctrl+R")}
            {alignBtn("justify", <TextAlignJustify20Regular />, "İki Yana Yasla", "Ctrl+J")}
            <RibbonMenuButton size="small" icon={<TextLineSpacing20Regular />} tip="Satır ve Paragraf Aralığı" menu={lineMenu} />
            <ColorSplitButton icon={<PaintBucket20Regular />} tip="Gölgelendirme" color={shade} noneLabel="Renk Yok" onColorChange={setShade} onApply={applyShading} />
            <RibbonSplitButton size="small" icon={<BorderBottom20Regular />} tip="Kenarlıklar" onClick={() => border("bottom")} menu={borderMenu} />
          </RibbonRow>
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Stiller" icon={<TextFontSize20Regular />} onLauncher={() => ctx.setView({ stylesPane: !ctx.view.stylesPane })} launcherTip="Stiller (Alt+Ctrl+Shift+S)">
        <RibbonGallery
          columns={6}
          itemWidth={76}
          itemHeight={58}
          expandedColumns={6}
          selectedId={fmt.charStyle ?? fmt.styleId}
          onSelect={(id) => applyStyle(editor, styles, id)}
          // Word gibi yeni oluşturulan stiller galerinin başında (en yenisi önce).
          items={[...Object.values(ctx.settings.customStyles ?? {}).filter((s) => s.gallery !== false && !GALLERY_STYLES.includes(s.id)).map((s) => s.id).reverse(), ...GALLERY_STYLES.filter((id) => styles[id])].filter((id) => styles[id]).map((id) => ({
            id,
            label: styles[id].name,
            preview: (
              <span className="bl-style-tile">
                <span className="bl-style-sample" style={previewStyle(styles, id)}>
                  AaBbÇçĞğ
                </span>
                <span className="bl-style-name">{styles[id].name}</span>
              </span>
            ),
          }))}
          footer={
            <>
              <MenuSeparator />
              <MenuItem icon={<TextEditStyle20Regular />} label="Stil Oluştur" onClick={() => ctx.openDialog("style")} />
              <MenuItem icon={<TextClearFormatting20Regular />} label="Biçimlendirmeyi Temizle" onClick={() => clearFormatting(editor)} />
              <MenuItem label="Stilleri Uygula..." onClick={() => ctx.setView({ stylesPane: true })} />
            </>
          }
        />
      </RibbonGroup>

      <RibbonGroup label="Düzenleme" icon={<Search20Regular />}>
        <RibbonStack>
          <RibbonSplitButton
            icon={<Search20Regular />}
            label="Bul"
            shortcut="Ctrl+F"
            onClick={ctx.openFind}
            menu={
              <Menu>
                <MenuItem icon={<Search20Regular />} label="Bul" shortcut="Ctrl+F" onClick={ctx.openFind} />
                <MenuItem label="Gelişmiş Bul..." onClick={() => ctx.openDialog("findReplace", "find")} />
                <MenuItem label="Git..." shortcut="Ctrl+G" onClick={() => ctx.openDialog("findReplace", "goto")} />
              </Menu>
            }
          />
          <RibbonButton icon={<ArrowSwap20Regular />} label="Değiştir" shortcut="Ctrl+H" onClick={() => ctx.openDialog("findReplace", "replace")} />
          <RibbonMenuButton
            icon={<Cursor20Regular />}
            label="Seç"
            menu={
              <Menu>
                <MenuItem label="Tümünü Seç" shortcut="Ctrl+A" onClick={() => chain().selectAll().run()} />
                <MenuItem label="Benzer Biçimlendirmeye Sahip Tüm Metni Seç" onClick={() => selectSimilar(editor)} />
              </Menu>
            }
          />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}

/** İmleçteki paragraf stiline sahip tüm paragrafları seçer (ilk ve son arası). */
function selectSimilar(editor: ReturnType<typeof useBelge>["editor"]) {
  const p = currentPara(editor);
  if (!p) return;
  const key = `${p.node.type.name}:${p.node.attrs.level ?? ""}:${p.node.attrs.styleId ?? ""}`;
  let first = -1;
  let last = -1;
  editor.state.doc.descendants((n, pos) => {
    if (n.type.name === "paragraph" || n.type.name === "heading") {
      if (`${n.type.name}:${n.attrs.level ?? ""}:${n.attrs.styleId ?? ""}` === key) {
        if (first < 0) first = pos + 1;
        last = pos + n.nodeSize - 1;
      }
      return false;
    }
    return true;
  });
  if (first >= 0) editor.chain().focus().setTextSelection({ from: first, to: last }).run();
}
