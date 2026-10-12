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
} from "@fluentui/react-icons";
import { RibbonButton, RibbonCheckbox, RibbonGroup, RibbonStack } from "@officehere/ui";
import { getUserName } from "@officehere/ui";
import { uid } from "../model/factory";
import { updateSlide } from "../model/ops";
import type { RibbonCtx } from "./types";

export function SlideShowTab({ ctx }: { ctx: RibbonCtx }) {
  const { state, actions: a, cmd } = ctx;
  const slide = state.pres.slides[state.slideIdx];
  const none = state.pres.slides.length === 0;
  return (
    <>
      <RibbonGroup label="Slayt Gösterisini Başlat">
        <RibbonButton size="large" icon={<Play24Regular />} label="Baştan Başla" shortcut="F5" disabled={none} description="Slayt gösterisini ilk slayttan başlatın." onClick={() => cmd.startShow("start")} />
        <RibbonButton size="large" icon={<SlidePlay24Regular />} label="Geçerli Slayttan" shortcut="Shift+F5" disabled={none} description="Slayt gösterisini bu slayttan başlatın." onClick={() => cmd.startShow("current")} />
        <RibbonButton size="large" icon={<Presenter24Regular />} label="Sunucu Görünümü" disabled={none} description="Notlarınızı, sonraki slaydı ve süreyi görerek sunun." onClick={cmd.presenterView} />
      </RibbonGroup>
      <RibbonGroup label="Ayarla">
        <RibbonButton size="large" icon={<SlideSettings24Regular />} label="Slayt Gösterisini Ayarla" description="Döngü, anlatım ve animasyon seçenekleri gibi gelişmiş ayarlar." onClick={() => ctx.openDialog("showSetup")} />
        <RibbonButton size="large" icon={<SlideHide24Regular />} label="Slaytı Gizle" checked={!!slide?.hidden} disabled={!slide} description="Seçili slaydı slayt gösterisi sırasında gizleyin." onClick={() => a.toggleHideSlides()} />
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
      <RibbonGroup label="Yazım Denetleme">
        <RibbonButton size="large" icon={<TextProofingTools24Regular />} label="Yazım Denetimi" checked={state.spellcheck} description="Yazım hatalarını kırmızı dalgalı çizgiyle işaretler (tarayıcı sözlüğü)." onClick={cmd.spelling} />
      </RibbonGroup>
      <RibbonGroup label="Açıklamalar">
        <RibbonButton
          size="large"
          icon={<CommentAdd24Regular />}
          label="Yeni Açıklama"
          disabled={!slide}
          onClick={() => {
            const sh = a.selectionBox();
            const c = { id: uid("c"), author: getUserName(), text: "", x: sh ? sh.x + sh.w : 20, y: sh ? sh.y : 20, date: Date.now() };
            a.store.commit((p) => updateSlide(p, state.slideIdx, (s) => ({ ...s, comments: [...(s.comments ?? []), c] })));
            a.store.set({ pane: "comments" });
          }}
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
      <RibbonGroup label="Sunu Görünümleri">
        <RibbonButton size="large" icon={<SlideText24Regular />} label="Normal" checked={state.view === "normal"} onClick={() => cmd.setView("normal")} />
        <RibbonButton size="large" icon={<SlideGrid24Regular />} label="Slayt Sıralayıcısı" checked={state.view === "sorter"} onClick={() => cmd.setView("sorter")} />
        <RibbonButton size="large" icon={<DocumentOnePage24Regular />} label="Not Sayfası" checked={state.view === "notes"} onClick={() => cmd.setView("notes")} />
        <RibbonButton size="large" icon={<ReadingModeMobile24Regular />} label="Okuma Görünümü" checked={state.view === "reading"} onClick={() => cmd.setView("reading")} />
      </RibbonGroup>
      <RibbonGroup label="Göster">
        <RibbonStack>
          <RibbonCheckbox label="Cetvel" checked={state.showRuler} onChange={(v) => a.store.set({ showRuler: v })} />
          <RibbonCheckbox label="Kılavuz Çizgileri" checked={state.showGrid} onChange={(v) => a.store.set({ showGrid: v })} />
          <RibbonCheckbox label="Kılavuzlar" checked={state.showGuides} onChange={(v) => a.store.set({ showGuides: v })} />
        </RibbonStack>
        <RibbonButton size="large" icon={<Notepad24Regular />} label="Notlar" checked={state.showNotes} onClick={() => a.store.set({ showNotes: !state.showNotes })} />
      </RibbonGroup>
      <RibbonGroup label="Yakınlaştır">
        <RibbonButton size="large" icon={<ZoomIn24Regular />} label="Yakınlaştır" onClick={() => ctx.openDialog("zoom")} />
        <RibbonButton size="large" icon={<ZoomFit24Regular />} label="Pencereye Sığdır" description="Sunuyu yakınlaştırarak slaytın pencereyi doldurmasını sağlayın." onClick={() => cmd.zoomTo(state.zoom, true)} />
      </RibbonGroup>
    </>
  );
}

export function HelpTab({ ctx }: { ctx: RibbonCtx }) {
  return (
    <>
      <RibbonGroup label="Yardım">
        <RibbonButton size="large" icon={<QuestionCircle24Regular />} label="Yardım" description="Klavye kısayolları ve temel kullanım." onClick={() => ctx.openDialog("shortcuts")} />
        <RibbonButton size="large" icon={<Keyboard24Regular />} label="Klavye Kısayolları" onClick={() => ctx.openDialog("shortcuts")} />
        <RibbonButton size="large" icon={<Info24Regular />} label="Hakkında" onClick={() => ctx.openDialog("about")} />
      </RibbonGroup>
    </>
  );
}
