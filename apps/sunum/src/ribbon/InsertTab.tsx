/** Ekle sekmesi: Slaytlar, Tablolar, Görüntüler, Çizimler, Metin, Simgeler. */
import {
  CalendarDate20Regular,
  DocumentHeaderFooter20Regular,
  Icons24Regular,
  Image32Regular,
  NumberSymbol20Regular,
  Shapes32Regular,
  SlideAdd32Regular,
  Symbols24Regular,
  Table32Regular,
  Textbox32Regular,
  TextEffects24Regular,
  Image24Regular,
  Shapes24Regular,
  SlideAdd24Regular,
  Table24Regular,
  Textbox24Regular,
  CommentAdd24Regular,
  Link24Regular,
  Link32Regular,
  ChartMultiple32Regular,
  Video24Regular,
  Video32Regular,
  Speaker232Regular, SlideSearch32Regular, CursorClick24Regular } from "@fluentui/react-icons";
import { getUserName, Menu, MenuItem, MenuSeparator, RibbonButton, RibbonGroup, RibbonMenuButton, RibbonSplitButton, RibbonStack, TableGridPicker, useMenuClose } from "@officehere/ui";
import { cssColor } from "../model/color";
import { LayoutGallery, ShapeGalleryMenu } from "./common";
import { WORDART_PRESETS } from "./presets";
import type { RibbonCtx } from "./types";
import { NewSlideExtras } from "./HomeTab";

function WordArtGrid({ ctx }: { ctx: RibbonCtx }) {
  const close = useMenuClose();
  const theme = ctx.state.pres.theme;
  return (
    <div className="oh-menu">
      <div className="preset-grid" style={{ gridTemplateColumns: "repeat(4, 56px)" }}>
        {WORDART_PRESETS.map((w, i) => (
          <button
            key={w.label}
            type="button"
            className="preset-btn"
            title={w.label}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              ctx.cmd.insertWordArt(i);
              close();
            }}
          >
            <span
              className="wordart-tile"
              style={{
                color: cssColor(w.color, theme),
                WebkitTextStroke: w.outline ? `1px ${cssColor(w.outline, theme)}` : undefined,
                textShadow: [w.shadow ? "2px 2px 2px rgba(0,0,0,.4)" : "", w.glow ? `0 0 6px ${cssColor(w.color, theme)}` : ""].filter(Boolean).join(",") || undefined,
                fontWeight: w.bold ? 800 : 600,
                fontFamily: "Georgia, serif",
              }}
            >
              A
            </span>
          </button>
        ))}
      </div>
    </div>
  );
}

export function InsertTab({ ctx }: { ctx: RibbonCtx }) {
  const { state, actions: a, cmd } = ctx;
  return (
    <>
      <RibbonGroup label="Slaytlar" icon={<SlideAdd24Regular />}>
        <RibbonSplitButton size="large" icon={<SlideAdd32Regular />} label="Yeni Slayt" shortcut="Ctrl+M" onClick={() => a.newSlide()} menu={<><LayoutGallery pres={state.pres} header={state.pres.theme.name || "Tema"} onPick={(id) => a.newSlide(id)} /><NewSlideExtras ctx={ctx} /></>} />
      </RibbonGroup>
      <RibbonGroup label="Tablolar" icon={<Table24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Table32Regular />}
          label="Tablo"
          description="Sunuya tablo ekleyin."
          menu={(close) => (
            <div className="oh-menu">
              <TableGridPicker onPick={(r, c) => a.insertTable(r, c)} />
              <MenuSeparator />
              <MenuItem label="Tablo Ekle..." onClick={() => (close(), ctx.openDialog("insertTable"))} />
            </div>
          )}
        />
      </RibbonGroup>
      <RibbonGroup label="Görüntüler" icon={<Image24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Image32Regular />}
          label="Resimler"
          description="Bilgisayarınızdan resim ekleyin."
          menu={
            <Menu>
              <MenuItem label="Bu Cihaz..." onClick={cmd.insertPictures} />
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Çizimler" icon={<Shapes24Regular />}>
        <RibbonMenuButton size="large" icon={<Shapes32Regular />} label="Şekiller" description="Hazır şekiller ekleyin: dikdörtgen, daire, ok, çizgi, belirtme çizgisi…" menu={<ShapeGalleryMenu onPick={cmd.pickShape} />} />
        <RibbonButton size="large" icon={<Icons24Regular />} label="Simgeler" description="Simge kitaplığından vektörel simge ekleyin." onClick={() => ctx.openDialog("icons")} />
        <RibbonButton size="large" icon={<ChartMultiple32Regular />} label="Grafik" description="Verileri göstermek için sütun, çizgi, pasta, çubuk, alan ya da halka grafiği ekleyin." onClick={cmd.insertChart} />
      </RibbonGroup>
      <RibbonGroup label="Bağlantılar" icon={<Link24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<SlideSearch32Regular />}
          label="Yakınlaştırma"
          description="Gösteride tıklanınca seçili slayda ya da bölüme yakınlaşarak giden canlı küçük resimler ekleyin."
          disabled={!state.pres.slides.length}
          menu={
            <Menu>
              <MenuItem label="Özet Yakınlaştırma..." onClick={() => ctx.openDialog("summaryZoom")} />
              <MenuItem label="Slayt Yakınlaştırma..." onClick={() => ctx.openDialog("slideZoom")} />
            </Menu>
          }
        />
        <RibbonButton size="large" icon={<Link32Regular />} label="Bağlantı" shortcut="Ctrl+K" description="Web sayfasına, e-posta adresine ya da sunudaki başka bir slayta köprü ekleyin." disabled={!state.pres.slides.length} onClick={cmd.insertLink} />
        <RibbonButton size="large" icon={<CursorClick24Regular />} label="Eylem" description="Seçili nesneye tıklanınca ya da üzerine gelinince yapılacak eylemi belirleyin." disabled={!state.sel.length} onClick={() => ctx.openDialog("action")} />
      </RibbonGroup>
      <RibbonGroup label="Açıklamalar" icon={<CommentAdd24Regular />}>
        <RibbonButton size="large" icon={<CommentAdd24Regular />} label="Açıklama" description="Seçili içerik hakkında açıklama ekleyin." disabled={!state.pres.slides.length} onClick={() => a.addComment(getUserName())} />
      </RibbonGroup>
      <RibbonGroup label="Metin" icon={<Textbox24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Textbox32Regular />}
          label="Metin Kutusu"
          description="Metin kutusu çizin; yazmaya başlamak için tıklayın."
          menu={
            <Menu>
              <MenuItem label="Yatay Metin Kutusu Çiz" onClick={() => a.store.set({ insert: { kind: "textbox" } })} />
              <MenuItem label="Dikey Metin Kutusu" onClick={() => a.store.set({ insert: { kind: "textbox", vertical: true } })} />
            </Menu>
          }
        />
        <RibbonButton size="large" icon={<DocumentHeaderFooter20Regular />} label="Üst Bilgi ve Alt Bilgi" description="Tarih/saat, slayt numarası ve alt bilgi metnini düzenleyin." onClick={() => ctx.openDialog("headerFooter")} />
        <RibbonMenuButton size="large" icon={<TextEffects24Regular />} label="Sanatsal Metin" description="Slaydınıza dekoratif metin ekleyin." menu={<WordArtGrid ctx={ctx} />} />
        <RibbonStack>
          <RibbonButton icon={<CalendarDate20Regular />} label="Tarih ve Saat" onClick={() => ctx.openDialog("dateTime")} />
          <RibbonButton icon={<NumberSymbol20Regular />} label="Slayt Numarası" onClick={cmd.insertSlideNumber} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Simgeler" icon={<Symbols24Regular />}>
        <RibbonButton size="large" icon={<Symbols24Regular />} label="Simge" description="Klavyede bulunmayan karakterler ekleyin." onClick={() => ctx.openDialog("symbol")} />
      </RibbonGroup>
      <RibbonGroup label="Medya" icon={<Video24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Video32Regular />}
          label="Video"
          description="Slayta video ekleyin."
          menu={
            <Menu>
              <MenuItem label="Bu Cihaz..." onClick={() => cmd.insertMedia("video")} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<Speaker232Regular />}
          label="Ses"
          description="Slayta ses ekleyin."
          menu={
            <Menu>
              <MenuItem label="Bilgisayarımdaki Ses..." onClick={() => cmd.insertMedia("audio")} />
            </Menu>
          }
        />
      </RibbonGroup>
    </>
  );
}
