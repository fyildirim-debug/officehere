/** Şerit sekmeleri (bağlamsal sekmeler seçime göre görünür). */
import type { RibbonTabDef } from "@officehere/ui";
import { MasterTab } from "../ribbon/MasterTab";
import { ChartDesignTab, ChartFormatTab, PlaybackTab } from "../ribbon/ChartTabs";
import { HomeTab } from "../ribbon/HomeTab";
import { InsertTab } from "../ribbon/InsertTab";
import { DrawTab } from "../ribbon/DrawTab";
import { DesignTab } from "../ribbon/DesignTab";
import { TransitionsTab } from "../ribbon/TransitionsTab";
import { AnimationsTab } from "../ribbon/AnimationsTab";
import { GrayscaleTab, HelpTab, ReviewTab, SlideShowTab, ViewTab } from "../ribbon/MiscTabs";
import { PictureFormatTab, ShapeFormatTab } from "../ribbon/FormatTabs";
import { TableDesignTab, TableLayoutTab } from "../ribbon/TableTabs";
import type { RibbonCtx } from "../ribbon/types";

export function buildRibbonTabs(ctx: RibbonCtx): RibbonTabDef[] {
  const { state, actions: a } = ctx;
  const shapes = a.selectedShapes();
  const one = shapes.length === 1 ? shapes[0] : undefined;
  const hasSpLike = shapes.some((s) => s.kind === "sp" || s.kind === "group");
  const hasPic = !!one && one.kind === "pic";
  const hasTable = !!one && one.kind === "table";
  const hasChart = !!one && one.kind === "chart";
  const hasMedia = !!one && one.kind === "pic" && !!one.media;
  const editView = state.view === "normal" || state.view === "outline" || state.view === "master";
  const masterView = state.view === "master";
  const gray = state.colorMode !== "color" && editView;
  return [
    ...(masterView ? [{ id: "master", label: "Asıl Slayt", content: <MasterTab ctx={ctx} /> }] : []),
    ...(gray ? [{ id: "grayscale", label: state.colorMode === "bw" ? "Siyah Beyaz" : "Gri Tonlama", content: <GrayscaleTab ctx={ctx} /> }] : []),
    { id: "home", label: "Giriş", content: <HomeTab ctx={ctx} /> },
    { id: "insert", label: "Ekle", content: <InsertTab ctx={ctx} /> },
    { id: "draw", label: "Çizim", content: <DrawTab ctx={ctx} /> },
    { id: "design", label: "Tasarım", hidden: masterView, content: <DesignTab ctx={ctx} /> },
    { id: "transitions", label: "Geçişler", hidden: masterView, content: <TransitionsTab ctx={ctx} /> },
    { id: "animations", label: "Animasyonlar", hidden: masterView, content: <AnimationsTab ctx={ctx} /> },
    { id: "slideshow", label: "Slayt Gösterisi", hidden: masterView, content: <SlideShowTab ctx={ctx} /> },
    { id: "review", label: "Gözden Geçir", content: <ReviewTab ctx={ctx} /> },
    { id: "view", label: "Görünüm", content: <ViewTab ctx={ctx} /> },
    { id: "help", label: "Yardım", content: <HelpTab ctx={ctx} /> },
    { id: "shapeFormat", label: "Şekil Biçimi", contextual: true, hidden: !hasSpLike || !editView, content: <ShapeFormatTab ctx={ctx} /> },
    { id: "pictureFormat", label: hasMedia ? (one?.kind === "pic" && one.media?.kind === "audio" ? "Ses Biçimi" : "Video Biçimi") : "Resim Biçimi", contextual: true, hidden: !hasPic || !editView, content: <PictureFormatTab ctx={ctx} /> },
    { id: "playback", label: "Kayıttan Yürüt", contextual: true, hidden: !hasMedia || !editView, content: <PlaybackTab ctx={ctx} /> },
    { id: "chartDesign", label: "Grafik Tasarımı", contextual: true, hidden: !hasChart || !editView, content: <ChartDesignTab ctx={ctx} /> },
    { id: "chartFormat", label: "Biçim", contextual: true, hidden: !hasChart || !editView, content: <ChartFormatTab ctx={ctx} /> },
    { id: "tableDesign", label: "Tablo Tasarımı", contextual: true, hidden: !hasTable || !editView, content: <TableDesignTab ctx={ctx} /> },
    { id: "tableLayout", label: "Düzen", contextual: true, hidden: !hasTable || !editView, content: <TableLayoutTab ctx={ctx} /> },
  ];
}
