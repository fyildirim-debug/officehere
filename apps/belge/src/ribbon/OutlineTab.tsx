// Anahat sekmesi (Görünüm › Anahat): Anahat Araçları (düzey, yükselt/düşür, taşı, genişlet/daralt, düzey göster)
// ve Anahat Görünümünü Kapat.
import {
  Add20Regular,
  ArrowDown20Regular,
  ArrowLeft20Regular,
  ArrowNext20Regular,
  ArrowPrevious20Regular,
  ArrowRight20Regular,
  ArrowUp20Regular,
  DismissCircle32Regular,
  Subtract20Regular,
  TextBulletListTree24Regular,
} from "@fluentui/react-icons";
import { RibbonButton, RibbonCheckbox, RibbonGroup, RibbonRow, RibbonStack } from "@officehere/ui";
import { useEditorState } from "@tiptap/react";
import { useBelge } from "../state/context";
import { moveBlock, outlineLevelOf, outlineState, setOutline, setOutlineLevel, shiftOutline, toggleCollapse } from "../editor/outline";

const LEVELS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
const levelLabel = (l: number) => (l === 10 ? "Gövde Metni" : `Düzey ${l}`);

export function OutlineTab() {
  const ctx = useBelge();
  const ed = ctx.mainEditor;
  // Anahat durumu ve imleçteki düzey her işlemde yeniden okunur (onay kutusu ve süzgeç durumu yansıtsın).
  const view = useEditorState({
    editor: ed,
    selector: ({ editor }) => {
      const $p = editor.state.selection.$from;
      const st = outlineState(editor.state);
      return { cur: $p.depth >= 2 ? outlineLevelOf($p.node(2)) : 10, level: st.level, firstLineOnly: st.firstLineOnly };
    },
    equalityFn: (a, b) => !!a && !!b && a.cur === b.cur && a.level === b.level && a.firstLineOnly === b.firstLineOnly,
  });
  const cur = view?.cur ?? 10;
  const s = { level: view?.level ?? 0, firstLineOnly: view?.firstLineOnly ?? false };
  return (
    <>
      <RibbonGroup label="Anahat Araçları" icon={<TextBulletListTree24Regular />}>
        <RibbonStack>
          <RibbonRow>
            <RibbonButton size="small" icon={<ArrowPrevious20Regular />} tip="Başlık 1'e Yükselt" onClick={() => setOutlineLevel(ed, 1)} />
            <RibbonButton size="small" icon={<ArrowLeft20Regular />} tip="Yükselt" shortcut="Alt+Shift+Sol" onClick={() => shiftOutline(ed, -1)} />
            <select className="bl-ol-select" aria-label="Anahat düzeyi" value={cur} onChange={(e) => setOutlineLevel(ed, Number(e.target.value))}>
              {LEVELS.map((l) => (
                <option key={l} value={l}>
                  {levelLabel(l)}
                </option>
              ))}
            </select>
            <RibbonButton size="small" icon={<ArrowRight20Regular />} tip="Düşür" shortcut="Alt+Shift+Sağ" onClick={() => shiftOutline(ed, 1)} />
            <RibbonButton size="small" icon={<ArrowNext20Regular />} tip="Gövde Metnine Düşür" onClick={() => setOutlineLevel(ed, 10)} />
          </RibbonRow>
          <RibbonRow>
            <RibbonButton size="small" icon={<ArrowUp20Regular />} tip="Yukarı Taşı" shortcut="Alt+Shift+Yukarı" onClick={() => moveBlock(ed, -1)} />
            <RibbonButton size="small" icon={<ArrowDown20Regular />} tip="Aşağı Taşı" shortcut="Alt+Shift+Aşağı" onClick={() => moveBlock(ed, 1)} />
            <RibbonButton size="small" icon={<Add20Regular />} tip="Genişlet" onClick={() => toggleCollapse(ed, false)} />
            <RibbonButton size="small" icon={<Subtract20Regular />} tip="Daralt" onClick={() => toggleCollapse(ed, true)} />
          </RibbonRow>
        </RibbonStack>
        <RibbonStack>
          <label className="bl-ol-show">
            Düzey Göster:
            <select aria-label="Düzey Göster" value={s.level} onChange={(e) => setOutline(ed, { level: Number(e.target.value) })}>
              <option value={0}>Tüm Düzeyler</option>
              {LEVELS.slice(0, 9).map((l) => (
                <option key={l} value={l}>
                  Düzey {l}
                </option>
              ))}
            </select>
          </label>
          <RibbonCheckbox label="Yalnızca İlk Satırı Göster" checked={s.firstLineOnly} onChange={(v) => setOutline(ed, { firstLineOnly: v })} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Kapat" icon={<DismissCircle32Regular />}>
        <RibbonButton size="large" icon={<DismissCircle32Regular />} label="Anahatı Kapat" tip="Anahat Görünümünü Kapat" onClick={() => ctx.setView({ mode: "print" })} />
      </RibbonGroup>
    </>
  );
}
