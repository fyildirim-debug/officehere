// Ekle sekmesi: Sayfalar, Tablolar, Çizimler, Bağlantılar, Açıklamalar, Üst Bilgi ve Alt Bilgi, Metin, Simgeler.
import {
  Bookmark20Regular,
  Calendar20Regular,
  CommentAdd32Regular,
  DocumentAdd24Regular,
  DocumentFooter24Regular,
  DocumentHeader32Regular,
  DocumentOnePage32Regular,
  DocumentPageBreak24Regular,
  DocumentPageNumber24Regular,
  Image32Regular,
  LineHorizontal120Regular,
  Link32Regular,
  Shapes32Regular,
  Symbols20Regular,
  Table32Regular,
  TextAddT20Regular,
  TableAdd20Regular,
  Delete20Regular,
  Edit20Regular,
} from "@fluentui/react-icons";
import { Menu, MenuHeader, MenuItem, MenuSeparator, RibbonButton, RibbonGroup, RibbonMenuButton, RibbonStack, TableGridPicker, useMenuClose } from "@officehere/ui";
import { useBelge } from "../state/context";
import { COMMON_SYMBOLS, COVER_PAGES, FOOTER_PRESETS, HEADER_PRESETS, pageNumberDoc, stripPageFields } from "../templates/parts";
import { SHAPES, shapeSvg, svgToPng, type ShapeKind } from "../editor/image";
import { insertTable } from "./tableUtil";

function SymbolGrid({ onPick }: { onPick: (s: string) => void }) {
  const close = useMenuClose();
  return (
    <div className="bl-symgrid">
      {COMMON_SYMBOLS.map((s) => (
        <button
          key={s}
          type="button"
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => {
            onPick(s);
            close();
          }}
        >
          {s}
        </button>
      ))}
    </div>
  );
}

function ShapeGrid({ onPick }: { onPick: (k: ShapeKind) => void }) {
  const close = useMenuClose();
  return (
    <div className="bl-shapegrid">
      {SHAPES.map((s) => (
        <button
          key={s.id}
          type="button"
          title={s.label}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => {
            onPick(s.id);
            close();
          }}
          dangerouslySetInnerHTML={{ __html: shapeSvg(s.id, "#DAE9F7", "#156082", 32, 24) }}
        />
      ))}
    </div>
  );
}

export function InsertTab() {
  const ctx = useBelge();
  const { editor, settings, updateSettings } = ctx;
  const chain = () => editor.chain().focus();

  const insertCover = (id: string) => {
    const def = COVER_PAGES.find((c) => c.id === id);
    if (!def) return;
    const title = settings.properties.title || "Belge Başlığı";
    const author = settings.properties.author || "Yazar Adı";
    const date = new Date().toLocaleDateString("tr-TR", { day: "numeric", month: "long", year: "numeric" });
    editor
      .chain()
      .focus()
      .insertContentAt(0, [...def.build(title, date, author), { type: "pageBreak" }])
      .setTextSelection(1)
      .run();
  };

  const insertShape = async (k: ShapeKind) => {
    const w = k === "line" || k === "arrow" ? 200 : 160;
    const h = k === "line" || k === "arrow" ? 60 : 110;
    const src = await svgToPng(shapeSvg(k, "#156082", "#0F4761", w, h), w, h);
    chain().setImage({ src, alt: SHAPES.find((s) => s.id === k)?.label, width: w, height: h }).run();
  };

  const setHeader = (json: typeof settings.header) => {
    updateSettings((s) => ({ ...s, header: json }));
    ctx.editHeaderFooter("header");
  };
  const setFooter = (json: typeof settings.footer) => {
    updateSettings((s) => ({ ...s, footer: json }));
    ctx.editHeaderFooter("footer");
  };

  return (
    <>
      <RibbonGroup label="Sayfalar" icon={<DocumentOnePage32Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<DocumentOnePage32Regular />}
          label="Kapak Sayfası"
          menu={
            <div className="oh-menu bl-cover-menu">
              <MenuHeader>Yerleşik</MenuHeader>
              <div className="bl-cover-grid">
                {COVER_PAGES.map((c) => (
                  <CoverTile key={c.id} id={c.id} label={c.label} onPick={insertCover} />
                ))}
              </div>
            </div>
          }
        />
        <RibbonStack>
          <RibbonButton icon={<DocumentAdd24Regular />} label="Boş Sayfa" onClick={() => chain().setPageBreak().setPageBreak().run()} />
          <RibbonButton icon={<DocumentPageBreak24Regular />} label="Sayfa Sonu" shortcut="Ctrl+Enter" onClick={() => chain().setPageBreak().run()} />
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Tablolar" icon={<Table32Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Table32Regular />}
          label="Tablo"
          menu={
            <div className="oh-menu">
              <TableGridPicker onPick={(r, c) => insertTable(editor, r, c)} />
              <MenuSeparator />
              <MenuItem icon={<TableAdd20Regular />} label="Tablo Ekle..." onClick={() => ctx.openDialog("insertTable")} />
            </div>
          }
        />
      </RibbonGroup>

      <RibbonGroup label="Çizimler" icon={<Image32Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Image32Regular />}
          label="Resimler"
          menu={
            <Menu>
              <MenuItem icon={<Image32Regular />} label="Bu Cihaz..." onClick={ctx.insertImageFromFile} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<Shapes32Regular />}
          label="Şekiller"
          menu={
            <div className="oh-menu">
              <MenuHeader>Temel Şekiller</MenuHeader>
              <ShapeGrid onPick={insertShape} />
            </div>
          }
        />
      </RibbonGroup>

      <RibbonGroup label="Bağlantılar" icon={<Link32Regular />}>
        <RibbonButton size="large" icon={<Link32Regular />} label="Bağlantı" shortcut="Ctrl+K" onClick={() => ctx.openDialog("link")} />
        <RibbonStack>
          <RibbonButton icon={<Bookmark20Regular />} label="Yer İşareti" onClick={() => ctx.openDialog("bookmark")} />
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Açıklamalar" icon={<CommentAdd32Regular />}>
        <RibbonButton size="large" icon={<CommentAdd32Regular />} label="Açıklama" shortcut="Ctrl+Alt+M" onClick={ctx.newComment} />
      </RibbonGroup>

      <RibbonGroup label="Üst Bilgi ve Alt Bilgi" icon={<DocumentHeader32Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<DocumentHeader32Regular />}
          label="Üst Bilgi"
          menu={
            <Menu width={240}>
              <MenuHeader>Yerleşik</MenuHeader>
              {HEADER_PRESETS.map((h) => (
                <MenuItem key={h.id} label={h.label} onClick={() => setHeader(h.content)} />
              ))}
              <MenuSeparator />
              <MenuItem icon={<Edit20Regular />} label="Üst Bilgiyi Düzenle" onClick={() => ctx.editHeaderFooter("header")} />
              <MenuItem icon={<Delete20Regular />} label="Üst Bilgiyi Kaldır" disabled={!settings.header} onClick={() => updateSettings((s) => ({ ...s, header: null }))} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<DocumentFooter24Regular />}
          label="Alt Bilgi"
          menu={
            <Menu width={240}>
              <MenuHeader>Yerleşik</MenuHeader>
              {FOOTER_PRESETS.map((h) => (
                <MenuItem key={h.id} label={h.label} onClick={() => setFooter(h.content)} />
              ))}
              <MenuSeparator />
              <MenuItem icon={<Edit20Regular />} label="Alt Bilgiyi Düzenle" onClick={() => ctx.editHeaderFooter("footer")} />
              <MenuItem icon={<Delete20Regular />} label="Alt Bilgiyi Kaldır" disabled={!settings.footer} onClick={() => updateSettings((s) => ({ ...s, footer: null }))} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<DocumentPageNumber24Regular />}
          label="Sayfa Numarası"
          menu={<PageNumberMenu />}
        />
      </RibbonGroup>

      <RibbonGroup label="Metin" icon={<TextAddT20Regular />}>
        <RibbonStack>
          <RibbonButton icon={<Calendar20Regular />} label="Tarih ve Saat" onClick={() => ctx.openDialog("dateTime")} />
          <RibbonButton icon={<LineHorizontal120Regular />} label="Yatay Çizgi" onClick={() => chain().setHorizontalRule().run()} />
        </RibbonStack>
      </RibbonGroup>

      <RibbonGroup label="Simgeler" icon={<Symbols20Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<span className="bl-omega">Ω</span>}
          label="Simge"
          menu={
            <div className="oh-menu">
              <SymbolGrid onPick={(s) => chain().insertContent(s).run()} />
              <MenuSeparator />
              <MenuItem icon={<Symbols20Regular />} label="Diğer Simgeler..." onClick={() => ctx.openDialog("symbol")} />
            </div>
          }
        />
      </RibbonGroup>
    </>
  );
}

function CoverTile({ id, label, onPick }: { id: string; label: string; onPick: (id: string) => void }) {
  const close = useMenuClose();
  return (
    <button
      type="button"
      className={`bl-cover-tile is-${id}`}
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => {
        onPick(id);
        close();
      }}
    >
      <span className="bl-cover-prev">
        <i />
        <b />
        <s />
      </span>
      <span>{label}</span>
    </button>
  );
}

function PageNumberMenu() {
  const ctx = useBelge();
  const set = (where: "header" | "footer", align: "left" | "center" | "right", total = false) => {
    ctx.updateSettings((s) => ({ ...s, [where]: pageNumberDoc(align, total) }));
  };
  return (
    <Menu width={220}>
      <MenuItem
        label="Sayfanın Üstü"
        submenu={
          <Menu>
            <MenuItem label="Düz Numara 1 (Solda)" onClick={() => set("header", "left")} />
            <MenuItem label="Düz Numara 2 (Ortada)" onClick={() => set("header", "center")} />
            <MenuItem label="Düz Numara 3 (Sağda)" onClick={() => set("header", "right")} />
            <MenuItem label="Sayfa X / Y" onClick={() => set("header", "right", true)} />
          </Menu>
        }
      />
      <MenuItem
        label="Sayfanın Altı"
        submenu={
          <Menu>
            <MenuItem label="Düz Numara 1 (Solda)" onClick={() => set("footer", "left")} />
            <MenuItem label="Düz Numara 2 (Ortada)" onClick={() => set("footer", "center")} />
            <MenuItem label="Düz Numara 3 (Sağda)" onClick={() => set("footer", "right")} />
            <MenuItem label="Sayfa X / Y" onClick={() => set("footer", "center", true)} />
          </Menu>
        }
      />
      <MenuItem
        label="Geçerli Konum"
        onClick={() => {
          if (ctx.headerFooterEditing) ctx.editor.chain().focus().insertPageField("PAGE").run();
          else ctx.notify("Sayfa numarası alanı üst bilgi ya da alt bilgi içine eklenebilir. Üst bilgiyi düzenlemek için sayfanın üst kenarına çift tıklayın.");
        }}
      />
      <MenuSeparator />
      <MenuItem label="Sayfa Numaralarını Kaldır" onClick={() => ctx.updateSettings((s) => ({ ...s, header: stripPageFields(s.header), footer: stripPageFields(s.footer) }))} />
    </Menu>
  );
}
