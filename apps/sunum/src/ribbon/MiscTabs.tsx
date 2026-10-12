/** Slayt Gösterisi, Gözden Geçir, Görünüm ve Yardım sekmeleri. */
import {
  CommentAdd24Regular,
  CommentDismiss20Regular,
  CommentArrowLeft20Regular,
  CommentArrowRight20Regular,
  CommentMultiple24Regular,
  DocumentOnePage24Regular,
  Info24Regular,
  Keyboard24Regular,
  Notepad24Regular,
  Presenter24Regular,
  QuestionCircle24Regular,
  ReadingModeMobile24Regular,
  SlideGrid24Regular,
  SlideHide24Regular,
  SlidePlay24Regular,
  SlideSettings24Regular,
  SlideText24Regular,
  TextProofingTools24Regular,
  ZoomFit24Regular,
  ZoomIn24Regular,
  Play24Regular,
  Comment24Regular,
  Eye24Regular,
  Settings24Regular,
  SlideMultiple24Regular,
  TextBulletListTree24Regular,
  SlideLayout24Regular,
  SlideMultipleArrowRight24Regular,
  Timer24Regular,
  Color24Regular,
  DarkTheme24Regular,
  CircleHalfFill24Regular,
} from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator, RibbonButton, RibbonCheckbox, RibbonGroup, RibbonMenuButton, RibbonStack } from "@officehere/ui";
import type { ShowSettings } from "../show/settings";
import { getUserName } from "@officehere/ui";
import { updateSlide } from "../model/ops";
import type { RibbonCtx } from "./types";

export function SlideShowTab({ ctx }: { ctx: RibbonCtx }) {
  const { state, actions: a, cmd } = ctx;
  const slide = state.pres.slides[state.slideIdx];
  const none = state.pres.slides.length === 0;
  const settings = ctx.show?.settings;
  const setSettings = (p: Partial<ShowSettings>) => settings && ctx.show?.set({ ...settings, ...p });
  const shows = state.pres.customShows ?? [];
  return (
    <>
      <RibbonGroup label="Slayt Gösterisini Başlat" icon={<Play24Regular />}>
        <RibbonButton size="large" icon={<Play24Regular />} label="Baştan" shortcut="F5" disabled={none} description="Slayt gösterisini ilk slayttan başlatın." onClick={() => cmd.startShow("start")} />
        <RibbonButton size="large" icon={<SlidePlay24Regular />} label="Geçerli Slayttan" shortcut="Shift+F5" disabled={none} description="Slayt gösterisini bu slayttan başlatın." onClick={() => cmd.startShow("current")} />
        <RibbonMenuButton
          size="large"
          icon={<SlideMultipleArrowRight24Regular />}
          label="Özel Slayt Gösterisi"
          disabled={none}
          description="Sunudaki slaytların yalnız bir bölümünü, istediğiniz sırayla gösterin."
          menu={
            <Menu>
              {shows.map((c) => (
                <MenuItem key={c.id} label={c.name} disabled={!c.slideIds.length} onClick={() => cmd.startCustomShow(c.id)} />
              ))}
              {shows.length > 0 && <MenuSeparator />}
              <MenuItem label="Özel Gösteriler..." onClick={cmd.customShows} />
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Ayarla" icon={<Settings24Regular />}>
        <RibbonButton size="large" icon={<SlideSettings24Regular />} label="Slayt Gösterisini Ayarla" tip="Slayt Gösterisini Ayarla" description="Gösteri türü, slayt aralığı, döngü, kalem rengi ve ilerletme gibi gelişmiş ayarlar." onClick={() => ctx.openDialog("showSetup")} />
        <RibbonButton size="large" icon={<SlideHide24Regular />} label="Slaytı Gizle" checked={!!slide?.hidden} disabled={!slide} description="Seçili slaydı slayt gösterisi sırasında gizleyin." onClick={() => a.toggleHideSlides()} />
        <RibbonButton size="large" icon={<Timer24Regular />} label="Zamanlama Provası" disabled={none} description="Gösteriyi tam ekranda prova edin; her slaytta geçen süre kaydedilir." onClick={cmd.rehearseTimings} />
        {settings && (
          <RibbonStack>
            <RibbonCheckbox label="Anlatımları Yürüt" checked={!settings.noNarration} onChange={(v) => setSettings({ noNarration: !v })} />
            <RibbonCheckbox label="Zamanlamaları Kullan" checked={settings.advance === "timings"} onChange={(v) => setSettings({ advance: v ? "timings" : "manual" })} />
          </RibbonStack>
        )}
      </RibbonGroup>
      <RibbonGroup label="Monitörler" icon={<Presenter24Regular />}>
        <RibbonButton size="large" icon={<Presenter24Regular />} label="Sunucu Görünümü" disabled={none} description="İzleyici gösterisi ikinci pencerede açılır; bu pencerede notlarınızı, sonraki slaydı ve süreyi görürsünüz." onClick={cmd.presenterView} />
        {settings && (
          <RibbonStack>
            <RibbonCheckbox label="Sunucu Görünümünü Kullan" checked={settings.presenter} onChange={(v) => setSettings({ presenter: v, type: v ? "speaker" : settings.type })} />
          </RibbonStack>
        )}
      </RibbonGroup>
    </>
  );
}

export function ReviewTab({ ctx }: { ctx: RibbonCtx }) {
  const { state, actions: a, cmd } = ctx;
  const slide = state.pres.slides[state.slideIdx];
  const comments = slide?.comments ?? [];
  const goComment = (dir: 1 | -1) => {
    const slides = state.pres.slides;
    for (let k = 1; k <= slides.length; k++) {
      const i = (state.slideIdx + dir * k + slides.length * 2) % slides.length;
      if (slides[i].comments?.length) {
        a.goToSlide(i);
        a.store.set({ pane: "comments" });
        return;
      }
    }
  };
  return (
    <>
      <RibbonGroup label="Yazım Denetleme" icon={<TextProofingTools24Regular />}>
        <RibbonButton size="large" icon={<TextProofingTools24Regular />} label="Yazım Denetimi" shortcut="F7" checked={state.pane === "spelling"} description="Yazım Denetimi bölmesini açar: denetim dili, açma/kapama ve metin kutularında sırayla gezinme." onClick={cmd.spelling} />
      </RibbonGroup>
      <RibbonGroup label="Açıklamalar" icon={<Comment24Regular />}>
        <RibbonButton
          size="large"
          icon={<CommentAdd24Regular />}
          label="Yeni Açıklama"
          disabled={!slide}
          onClick={() => a.addComment(getUserName())}
        />
        <RibbonStack>
          <RibbonButton icon={<CommentDismiss20Regular />} label="Sil" disabled={!comments.length} onClick={() => a.store.commit((p) => updateSlide(p, state.slideIdx, (s) => ({ ...s, comments: (s.comments ?? []).slice(0, -1) })))} />
          <RibbonButton icon={<CommentArrowLeft20Regular />} label="Önceki" onClick={() => goComment(-1)} />
          <RibbonButton icon={<CommentArrowRight20Regular />} label="Sonraki" onClick={() => goComment(1)} />
        </RibbonStack>
        <RibbonButton size="large" icon={<CommentMultiple24Regular />} label="Açıklamaları Göster" checked={state.pane === "comments"} onClick={() => cmd.togglePane("comments")} />
      </RibbonGroup>
    </>
  );
}

export function ViewTab({ ctx }: { ctx: RibbonCtx }) {
  const { state, actions: a, cmd } = ctx;
  return (
    <>
      <RibbonGroup label="Sunu Görünümleri" icon={<SlideMultiple24Regular />}>
        <RibbonButton size="large" icon={<SlideText24Regular />} label="Normal" checked={state.view === "normal"} onClick={() => cmd.setView("normal")} />
        <RibbonButton size="large" icon={<TextBulletListTree24Regular />} label="Ana Hat Görünümü" checked={state.view === "outline"} description="Slayt başlıklarını ve metinlerini ana hat olarak düzenleyin." onClick={() => cmd.setView("outline")} />
        <RibbonButton size="large" icon={<SlideGrid24Regular />} label="Slayt Sıralayıcısı" checked={state.view === "sorter"} onClick={() => cmd.setView("sorter")} />
        <RibbonButton size="large" icon={<DocumentOnePage24Regular />} label="Not Sayfası" checked={state.view === "notes"} onClick={() => cmd.setView("notes")} />
        <RibbonButton size="large" icon={<ReadingModeMobile24Regular />} label="Okuma Görünümü" checked={state.view === "reading"} onClick={() => cmd.setView("reading")} />
      </RibbonGroup>
      <RibbonGroup label="Asıl Görünümler" icon={<SlideLayout24Regular />}>
        <RibbonButton size="large" icon={<SlideLayout24Regular />} label="Asıl Slayt" checked={state.view === "master"} description="Asıl slaydı ve düzenlerini düzenleyin: yer tutucu konumları, yazı tipleri ve arka plan tüm slaytlara yansır." disabled={!state.pres.layouts.length} onClick={() => cmd.setView(state.view === "master" ? "normal" : "master")} />
      </RibbonGroup>
      <RibbonGroup label="Renk/Gri Tonlama" icon={<Color24Regular />}>
        <RibbonButton size="large" icon={<Color24Regular />} label="Renkli" checked={state.colorMode === "color"} onClick={() => a.store.set({ colorMode: "color" })} />
        <RibbonButton size="large" icon={<DarkTheme24Regular />} label="Gri Tonlama" checked={state.colorMode === "gray"} description="Sunuyu gri tonlamalı olarak görüntüleyin (ör. siyah beyaz yazdırma öncesi)." onClick={() => a.store.set({ colorMode: "gray", ribbonTab: "grayscale" })} />
        <RibbonButton size="large" icon={<CircleHalfFill24Regular />} label="Siyah Beyaz" checked={state.colorMode === "bw"} description="Sunuyu yalnız siyah ve beyazla görüntüleyin." onClick={() => a.store.set({ colorMode: "bw", ribbonTab: "grayscale" })} />
      </RibbonGroup>
      <RibbonGroup label="Göster" icon={<Eye24Regular />}>
        <RibbonStack>
          <RibbonCheckbox label="Cetvel" checked={state.showRuler} onChange={(v) => a.store.set({ showRuler: v })} />
          <RibbonCheckbox label="Kılavuz Çizgileri" checked={state.showGrid} onChange={(v) => a.store.set({ showGrid: v })} />
          <RibbonCheckbox label="Kılavuzlar" checked={state.showGuides} onChange={(v) => a.store.set({ showGuides: v })} />
        </RibbonStack>
        <RibbonButton size="large" icon={<Notepad24Regular />} label="Notlar" checked={state.showNotes} onClick={() => a.store.set({ showNotes: !state.showNotes })} />
      </RibbonGroup>
      <RibbonGroup label="Yakınlaştır" icon={<ZoomIn24Regular />}>
        <RibbonButton size="large" icon={<ZoomIn24Regular />} label="Yakınlaştır" onClick={() => ctx.openDialog("zoom")} />
        <RibbonButton size="large" icon={<ZoomFit24Regular />} label="Pencereye Sığdır" description="Sunuyu yakınlaştırarak slaytın pencereyi doldurmasını sağlayın." onClick={() => cmd.zoomTo(state.zoom, true)} />
      </RibbonGroup>
    </>
  );
}

/** Gri Tonlama / Siyah Beyaz görünümünün bağlamsal sekmesi. */
export function GrayscaleTab({ ctx }: { ctx: RibbonCtx }) {
  const { state, actions: a } = ctx;
  return (
    <>
      <RibbonGroup label="Görünüm" icon={<DarkTheme24Regular />}>
        <RibbonButton size="large" icon={<DarkTheme24Regular />} label="Gri Tonlama" checked={state.colorMode === "gray"} onClick={() => a.store.set({ colorMode: "gray" })} />
        <RibbonButton size="large" icon={<CircleHalfFill24Regular />} label="Siyah Beyaz" checked={state.colorMode === "bw"} onClick={() => a.store.set({ colorMode: "bw" })} />
      </RibbonGroup>
      <RibbonGroup label="Kapat" icon={<Color24Regular />}>
        <RibbonButton size="large" icon={<Color24Regular />} label="Renkli Görünüme Dön" onClick={() => a.store.set({ colorMode: "color", ribbonTab: "view" })} />
      </RibbonGroup>
    </>
  );
}

export function HelpTab({ ctx }: { ctx: RibbonCtx }) {
  return (
    <>
      <RibbonGroup label="Yardım" icon={<QuestionCircle24Regular />}>
        <RibbonButton size="large" icon={<QuestionCircle24Regular />} label="Yardım" description="Klavye kısayolları ve temel kullanım." onClick={() => ctx.openDialog("shortcuts")} />
        <RibbonButton size="large" icon={<Keyboard24Regular />} label="Klavye Kısayolları" onClick={() => ctx.openDialog("shortcuts")} />
        <RibbonButton size="large" icon={<Info24Regular />} label="Hakkında" onClick={() => ctx.openDialog("about")} />
      </RibbonGroup>
    </>
  );
}
