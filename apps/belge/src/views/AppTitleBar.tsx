// Başlık çubuğu: Hızlı Erişim Araç Çubuğu (Kaydet, Geri Al, Yinele, Baskı Önizleme ve Yazdır) ve arama.
import { ArrowRedo20Regular, ArrowUndo20Regular, Print20Regular, Save20Regular } from "@fluentui/react-icons";
import { QatButton, TitleBar } from "@officehere/ui";
import type { Editor } from "@tiptap/core";
import { repeatOrRedo } from "../editor/repeat";

interface Props {
  title: string;
  dirty: boolean;
  active: Editor;
  canUndo: boolean;
  canRedo: boolean;
  onSave: () => void;
  onPrint: () => void;
  onFind: () => void;
}

export function AppTitleBar({ title, dirty, active, canUndo, canRedo, onSave, onPrint, onFind }: Props) {
  return (
    <TitleBar
      app="belge"
      title={title}
      dirty={dirty}
      onSearch={(q) => {
        onFind();
        requestAnimationFrame(() => {
          const input = document.querySelector<HTMLInputElement>(".bl-nav-search input");
          if (input) {
            const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
            setter?.call(input, q);
            input.dispatchEvent(new Event("input", { bubbles: true }));
          }
        });
      }}
      searchPlaceholder="Ara (Alt+Q)"
      quickAccess={
        <>
          <QatButton icon={<Save20Regular />} tip="Kaydet" shortcut="Ctrl+S" onClick={onSave} />
          <QatButton icon={<ArrowUndo20Regular />} tip="Geri Al" shortcut="Ctrl+Z" disabled={!canUndo} onClick={() => active.chain().focus().undo().run()} />
          <QatButton
            icon={<ArrowRedo20Regular />}
            tip="Yinele"
            shortcut="Ctrl+Y"
            disabled={!canRedo}
            onClick={() => {
              active.commands.focus();
              repeatOrRedo(active);
            }}
          />
          <QatButton icon={<Print20Regular />} tip="Baskı Önizleme ve Yazdır" shortcut="Ctrl+P" onClick={onPrint} />
        </>
      }
    />
  );
}
