// Gözden Geçir sekmesi: Yazım Denetleme, Dil, Açıklamalar, İzleme, Değişiklikler.
import { TextSelection } from "@tiptap/pm/state";
import {
  ArrowNext20Regular,
  ArrowPrevious20Regular,
  CheckmarkCircle32Regular,
  CommentAdd32Regular,
  CommentDismiss20Regular,
  CommentMultiple20Regular,
  DismissCircle32Regular,
  DocumentEdit24Regular,
  LocalLanguage24Regular,
  TextGrammarCheckmark24Regular,
  TextWordCount20Regular,
  TextEditStyle20Regular,
  Eye20Regular,
} from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator, RibbonButton, RibbonGroup, RibbonMenuButton, RibbonSplitButton, RibbonStack } from "@officehere/ui";
import { useBelge } from "../state/context";
import { commentAnchors } from "../views/CommentsPane";
import { listChanges, resolveChanges } from "../editor/trackChanges";

const MARKUP_LABEL = { simple: "Basit İşaretleme", all: "Tüm İşaretlemeler", none: "İşaretleme Yok", original: "Özgün" } as const;

export function ReviewTab({ lang, setLang, deleteComment, deleteAllComments }: { lang: string; setLang: (l: string) => void; deleteComment: (id: string) => void; deleteAllComments: () => void }) {
  const ctx = useBelge();
  const { editor } = ctx;
  const anchors = commentAnchors(editor);
  const pos = editor.state.selection.from;
  const current = anchors.find((a) => pos >= a.from && pos <= a.to);
  const changes = listChanges(editor.state.doc);
  const markup = ctx.view.markup ?? "all";
  const balloons = ctx.view.balloons ?? "revisions";

  const gotoComment = (dir: 1 | -1) => {
    if (!anchors.length) return;
    const list = dir > 0 ? anchors.find((a) => a.from > pos) ?? anchors[0] : [...anchors].reverse().find((a) => a.to < pos) ?? anchors[anchors.length - 1];
    editor.chain().focus().setTextSelection({ from: list.from, to: list.to }).scrollIntoView().run();
  };

  const gotoChange = (dir: 1 | -1) => {
    if (!changes.length) return;
    const c = dir > 0 ? changes.find((x) => x.from > pos) ?? changes[0] : [...changes].reverse().find((x) => x.to < pos) ?? changes[changes.length - 1];
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, c.from, c.to)).scrollIntoView());
    editor.commands.focus();
  };

  const resolve = (accept: boolean, all: boolean) => {
    const { state } = editor;
    let from = state.selection.from;
    let to = state.selection.to;
    if (all) {
      from = 0;
      to = state.doc.content.size;
    } else if (from === to) {
      const c = changes.find((x) => from >= x.from && from <= x.to);
      if (!c) return gotoChange(1);
      from = c.from;
      to = c.to;
    }
    editor.view.dispatch(resolveChanges(state, from, to, accept));
    if (!all) gotoChange(1);
  };

  return (
    <>
      <RibbonGroup label="Yazım Denetleme" icon={<TextGrammarCheckmark24Regular />}>
        <RibbonButton size="large" icon={<TextGrammarCheckmark24Regular />} label="Yazım ve Dil Bilgisi" shortcut="F7" description="Düzenleyici bölmesini açarak belgedeki yazım ve dil bilgisi önerilerini gözden geçirin." onClick={() => ctx.setView({ proofPane: true })} />
        <RibbonStack>
          <RibbonButton icon={<TextWordCount20Regular />} label="Sözcük Sayısı" shortcut="Ctrl+Shift+G" onClick={() => ctx.openDialog("wordCount")} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Dil" icon={<LocalLanguage24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<LocalLanguage24Regular />}
          label="Dil"
          menu={
            <Menu width={240}>
              <MenuItem label="Türkçe (Türkiye)" checked={lang === "tr"} onClick={() => setLang("tr")} />
              <MenuItem label="İngilizce (ABD)" checked={lang === "en-US"} onClick={() => setLang("en-US")} />
              <MenuItem label="Almanca (Almanya)" checked={lang === "de"} onClick={() => setLang("de")} />
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Açıklamalar" icon={<CommentAdd32Regular />}>
        <RibbonButton size="large" icon={<CommentAdd32Regular />} label="Yeni Açıklama" shortcut="Ctrl+Alt+M" onClick={ctx.newComment} />
        <RibbonStack>
          <RibbonSplitButton
            icon={<CommentDismiss20Regular />}
            label="Sil"
            disabled={!anchors.length}
            onClick={() => current && deleteComment(current.id)}
            menu={
              <Menu>
                <MenuItem label="Sil" disabled={!current} onClick={() => current && deleteComment(current.id)} />
                <MenuItem label="Belgedeki Tüm Açıklamaları Sil" onClick={deleteAllComments} />
              </Menu>
            }
          />
          <RibbonButton icon={<ArrowPrevious20Regular />} label="Önceki" disabled={!anchors.length} onClick={() => gotoComment(-1)} />
          <RibbonButton icon={<ArrowNext20Regular />} label="Sonraki" disabled={!anchors.length} onClick={() => gotoComment(1)} />
        </RibbonStack>
        <RibbonStack>
          <RibbonButton icon={<CommentMultiple20Regular />} label="Açıklamaları Göster" checked={ctx.view.showComments} onClick={() => ctx.setView({ showComments: !ctx.view.showComments })} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="İzleme" icon={<DocumentEdit24Regular />}>
        <RibbonButton size="large" icon={<DocumentEdit24Regular />} label="Değişiklikleri İzle" shortcut="Ctrl+Shift+E" checked={ctx.tracking} description="Bu belgede yapılan tüm değişiklikleri izleyin." onClick={ctx.toggleTrack} />
        <RibbonStack>
          <RibbonMenuButton
            icon={<TextEditStyle20Regular />}
            label={MARKUP_LABEL[markup]}
            tip="Görüntülenecek işaretleme"
            menu={
              <Menu width={220}>
                {(Object.keys(MARKUP_LABEL) as (keyof typeof MARKUP_LABEL)[]).map((k) => (
                  <MenuItem key={k} label={MARKUP_LABEL[k]} checked={markup === k} onClick={() => ctx.setView({ markup: k })} />
                ))}
              </Menu>
            }
          />
          <RibbonMenuButton
            icon={<Eye20Regular />}
            label="İşaretlemeyi Göster"
            menu={
              <Menu width={300}>
                <MenuItem label="Açıklamalar" checked={ctx.view.showComments} onClick={() => ctx.setView({ showComments: !ctx.view.showComments })} />
                <MenuItem
                  label="Balonlar"
                  submenu={
                    <Menu width={320}>
                      <MenuItem label="Düzeltmeleri Balonlarda Göster" checked={balloons === "revisions"} onClick={() => ctx.setView({ balloons: "revisions" })} />
                      <MenuItem label="Tüm Düzeltmeleri Satır İçinde Göster" checked={balloons === "inline"} onClick={() => ctx.setView({ balloons: "inline" })} />
                      <MenuItem label="Yalnızca Açıklamaları ve Biçimlendirmeyi Balonlarda Göster" checked={balloons === "formatting"} onClick={() => ctx.setView({ balloons: "formatting" })} />
                    </Menu>
                  }
                />
              </Menu>
            }
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Değişiklikler" icon={<CheckmarkCircle32Regular />}>
        <RibbonSplitButton
          size="large"
          icon={<CheckmarkCircle32Regular />}
          label="Kabul Et"
          disabled={!changes.length}
          onClick={() => resolve(true, false)}
          menu={
            <Menu>
              <MenuItem label="Kabul Et ve Sonrakine Geç" onClick={() => resolve(true, false)} />
              <MenuSeparator />
              <MenuItem label="Tüm Değişiklikleri Kabul Et" onClick={() => resolve(true, true)} />
            </Menu>
          }
        />
        <RibbonSplitButton
          size="large"
          icon={<DismissCircle32Regular />}
          label="Reddet"
          disabled={!changes.length}
          onClick={() => resolve(false, false)}
          menu={
            <Menu>
              <MenuItem label="Reddet ve Sonrakine Geç" onClick={() => resolve(false, false)} />
              <MenuSeparator />
              <MenuItem label="Tüm Değişiklikleri Reddet" onClick={() => resolve(false, true)} />
            </Menu>
          }
        />
        <RibbonStack>
          <RibbonButton icon={<ArrowPrevious20Regular />} label="Önceki" disabled={!changes.length} onClick={() => gotoChange(-1)} />
          <RibbonButton icon={<ArrowNext20Regular />} label="Sonraki" disabled={!changes.length} onClick={() => gotoChange(1)} />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}
