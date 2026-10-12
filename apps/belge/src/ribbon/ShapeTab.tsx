// Bağlamsal "Şekil Biçimi" sekmesi: metin kutusunun dolgusu, ana hattı ve boyutu.
import { NodeSelection } from "@tiptap/pm/state";
import { LineThickness20Regular, PaintBucket20Regular, Shapes24Regular, BorderOutside20Regular } from "@fluentui/react-icons";
import { ColorPalette, Menu, MenuItem, MenuSeparator, RibbonGroup, RibbonMenuButton, RibbonSpinner, RibbonStack } from "@officehere/ui";
import { useBelge } from "../state/context";
import { textBoxAt } from "../editor/formatState";

const CM = 96 / 2.54;
const WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];

export function ShapeTab() {
  const { editor } = useBelge();
  const at = textBoxAt(editor.state);
  if (!at) return null;
  const a = at.node.attrs as { w: number; h: number; border: string | null; borderWidth: number; fill: string | null };
  const update = (attrs: Record<string, unknown>) => {
    const tr = editor.state.tr.setNodeMarkup(at.pos, undefined, { ...at.node.attrs, ...attrs });
    // Kutu seçili kalır (biçimleme sürdürülebilsin).
    tr.setSelection(NodeSelection.create(tr.doc, at.pos));
    editor.view.dispatch(tr);
    editor.view.focus();
  };
  return (
    <>
      <RibbonGroup label="Şekil Stilleri" icon={<Shapes24Regular />}>
        <RibbonStack>
          <RibbonMenuButton
            icon={<PaintBucket20Regular />}
            label="Şekil Dolgusu"
            menu={
              <div className="oh-menu oh-menu-palette">
                <ColorPalette noneLabel="Dolgu Yok" value={a.fill ?? undefined} onSelect={(c) => update({ fill: c })} />
              </div>
            }
          />
          <RibbonMenuButton
            icon={<BorderOutside20Regular />}
            label="Şekil Ana Hattı"
            menu={
              <div className="oh-menu oh-menu-palette">
                <ColorPalette noneLabel="Ana Hat Yok" value={a.border ?? undefined} onSelect={(c) => update({ border: c })} />
              </div>
            }
          />
          <RibbonMenuButton
            icon={<LineThickness20Regular />}
            label="Kalınlık"
            menu={
              <Menu>
                {WEIGHTS.map((w) => (
                  <MenuItem key={w} label={`${w.toLocaleString("tr-TR")} nk`} checked={!!a.border && Math.abs(a.borderWidth - w) < 0.01} onClick={() => update({ borderWidth: w, border: a.border ?? "#000000" })} />
                ))}
                <MenuSeparator />
                <MenuItem label="Ana Hat Yok" onClick={() => update({ border: null })} />
              </Menu>
            }
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Boyut">
        <RibbonStack>
          <RibbonSpinner label="Yükseklik:" value={a.h / CM} step={0.1} min={0.3} max={60} unit="cm" decimals={2} width={78} onChange={(v) => update({ h: Math.round(v * CM) })} />
          <RibbonSpinner label="Genişlik:" value={a.w / CM} step={0.1} min={0.3} max={60} unit="cm" decimals={2} width={78} onChange={(v) => update({ w: Math.round(v * CM) })} />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}
