/** Ekle sekmesi: Tablolar, Çizimler, Grafikler, Mini Grafikler, Bağlantılar, Metin, Simgeler. */
import {
  DataBarVertical24Regular,
  DataLine24Regular,
  DataPie24Regular,
  DataScatter24Regular,
  DocumentHeaderFooter24Regular,
  ImageAdd24Regular,
  Link24Regular,
  NoteAdd24Regular,
  Shapes24Regular,
  Symbols24Regular,
  Table24Regular,
  Textbox24Regular,
  ChartMultiple24Regular,
} from "@fluentui/react-icons";
import { Menu, MenuHeader, MenuItem, RibbonButton, RibbonGroup, RibbonMenuButton } from "@officehere/ui";
import type { Controller } from "../controller";
import { insertChart, insertShape, pickAndInsertImage, SHAPE_KINDS } from "../actions/objects";
import { CHART_FAMILIES, ChartTypeIcon } from "../dialogs/ChartDialogs";
import type { ShapeDef } from "../model/types";
import { PivotGlyph, SparkGlyph } from "./glyphs";

function ChartMenu({ ctl, families }: { ctl: Controller; families: string[] }) {
  return (
    <Menu width={230}>
      {CHART_FAMILIES.filter((f) => families.includes(f.id)).map((f) => (
        <div key={f.id}>
          <MenuHeader>{f.label}</MenuHeader>
          {f.types.map((t) => (
            <MenuItem key={t.type} icon={<ChartTypeIcon type={t.type} size={20} />} label={t.label} onClick={() => insertChart(ctl, t.type)} />
          ))}
        </div>
      ))}
      <MenuItem label="Diğer Grafikler..." onClick={() => ctl.openDialog("chartType", { insert: true, type: CHART_FAMILIES.find((f) => families.includes(f.id))?.types[0].type })} />
    </Menu>
  );
}

export function ShapeIcon({ kind }: { kind: ShapeDef["kind"] }) {
  const s = { fill: "none", stroke: "currentColor", strokeWidth: 1.3 };
  return (
    <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true">
      {kind === "rect" && <rect x="3" y="5" width="14" height="10" {...s} />}
      {kind === "roundRect" && <rect x="3" y="5" width="14" height="10" rx="3" {...s} />}
      {kind === "ellipse" && <ellipse cx="10" cy="10" rx="7" ry="5" {...s} />}
      {kind === "triangle" && <path d="M10 4 L17 16 H3 Z" {...s} />}
      {kind === "arrow" && <path d="M3 8 H11 V5 L17 10 L11 15 V12 H3 Z" {...s} />}
      {kind === "line" && <path d="M3 16 L17 4" {...s} />}
      {kind === "textbox" && (
        <>
          <rect x="3" y="4" width="14" height="12" {...s} strokeDasharray="2 1" />
          <path d="M7 8 H13 M10 8 V13" {...s} />
        </>
      )}
    </svg>
  );
}

export function InsertTab({ ctl }: { ctl: Controller }) {
  return (
    <>
      <RibbonGroup label="Tablolar" icon={<Table24Regular />}>
        <RibbonButton size="large" icon={<PivotGlyph />} label="Özet Tablo" description="Karmaşık verileri özet tabloyla kolayca düzenleyip özetleyin." onClick={() => ctl.openDialog("createPivot")} />
        <RibbonButton size="large" icon={<Table24Regular />} label="Tablo" shortcut="Ctrl+T" description="Verileri düzenlemek ve çözümlemek için tablo oluşturur." onClick={() => ctl.openDialog("createTable")} />
      </RibbonGroup>
      <RibbonGroup label="Çizimler" icon={<ImageAdd24Regular />}>
        <RibbonButton size="large" icon={<ImageAdd24Regular />} label="Resimler" description="Bilgisayarınızdan resim ekler." onClick={() => pickAndInsertImage(ctl)} />
        <RibbonMenuButton
          size="large"
          icon={<Shapes24Regular />}
          label="Şekiller"
          menu={
            <Menu width={220}>
              <MenuHeader>Temel Şekiller</MenuHeader>
              {SHAPE_KINDS.map((k) => (
                <MenuItem key={k.kind} icon={<ShapeIcon kind={k.kind} />} label={k.label} onClick={() => insertShape(ctl, k.kind)} />
              ))}
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Grafikler" icon={<ChartMultiple24Regular />} onLauncher={() => ctl.openDialog("chartType", { insert: true })} launcherTip="Grafik Ekle">
        <RibbonButton size="large" icon={<ChartMultiple24Regular />} label="Önerilen Grafikler" description="Verileriniz için grafik türü seçin ve önizleyin." onClick={() => ctl.openDialog("chartType", { insert: true })} />
        <RibbonMenuButton size="large" icon={<DataBarVertical24Regular />} label="Sütun veya Çubuk" menu={<ChartMenu ctl={ctl} families={["column", "bar"]} />} />
        <RibbonMenuButton size="large" icon={<DataLine24Regular />} label="Çizgi veya Alan" menu={<ChartMenu ctl={ctl} families={["line", "area"]} />} />
        <RibbonMenuButton size="large" icon={<DataPie24Regular />} label="Pasta veya Halka" menu={<ChartMenu ctl={ctl} families={["pie"]} />} />
        <RibbonMenuButton size="large" icon={<DataScatter24Regular />} label="Dağılım" menu={<ChartMenu ctl={ctl} families={["scatter"]} />} />
      </RibbonGroup>
      <RibbonGroup label="Mini Grafikler" icon={<SparkGlyph kind="line" />}>
        <RibbonButton size="large" icon={<SparkGlyph kind="line" />} label="Çizgi" description="Tek hücreye çizgi mini grafik ekler." onClick={() => ctl.openDialog("sparkline", { type: "line" })} />
        <RibbonButton size="large" icon={<SparkGlyph kind="column" />} label="Sütun" description="Tek hücreye sütun mini grafik ekler." onClick={() => ctl.openDialog("sparkline", { type: "column" })} />
        <RibbonButton size="large" icon={<SparkGlyph kind="winLoss" />} label="Kazanç/Kayıp" description="Tek hücreye kazanç/kayıp mini grafiği ekler." onClick={() => ctl.openDialog("sparkline", { type: "winLoss" })} />
      </RibbonGroup>
      <RibbonGroup label="Bağlantılar" icon={<Link24Regular />}>
        <RibbonButton size="large" icon={<Link24Regular />} label="Bağlantı" shortcut="Ctrl+K" description="Web sayfasına ya da bu belgedeki bir yere bağlantı ekler." onClick={() => ctl.openDialog("hyperlink")} />
      </RibbonGroup>
      <RibbonGroup label="Açıklamalar" icon={<NoteAdd24Regular />}>
        <RibbonButton size="large" icon={<NoteAdd24Regular />} label="Not" shortcut="Shift+F2" description="Seçili hücreye not ekler." onClick={() => ctl.openDialog("note")} />
      </RibbonGroup>
      <RibbonGroup label="Metin" icon={<Textbox24Regular />}>
        <RibbonButton size="large" icon={<Textbox24Regular />} label="Metin Kutusu" onClick={() => insertShape(ctl, "textbox")} />
        <RibbonButton size="large" icon={<DocumentHeaderFooter24Regular />} label="Üst Bilgi ve Alt Bilgi" description="Yazdırılan sayfaların üstünde ve altında görünecek metni düzenler." onClick={() => ctl.openDialog("pageSetup", { tab: "hf" })} />
      </RibbonGroup>
      <RibbonGroup label="Simgeler" icon={<Symbols24Regular />}>
        <RibbonButton size="large" icon={<Symbols24Regular />} label="Simge" description="Klavyede bulunmayan karakterleri ekler." onClick={() => ctl.openDialog("symbol")} />
      </RibbonGroup>
    </>
  );
}
