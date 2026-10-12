// Madde işareti / numaralandırma / çok düzeyli liste kitaplıkları.
import type { Editor } from "@tiptap/core";
import { MenuHeader, MenuItem, MenuSeparator, useMenuClose } from "@officehere/ui";
import { BULLET_STYLES, MULTILEVEL_STYLES, NUMBER_STYLES } from "../editor/nodes";
import { toggleList } from "../editor/listApply";

function listNodeAttrs(editor: Editor, type: "bulletList" | "orderedList", attrs: Record<string, unknown>) {
  const { state } = editor;
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d);
    if (n.type.name === type) {
      // En dıştaki aynı türden listeye değil, imlecin içinde olduğu listeye uygula.
      editor.view.dispatch(state.tr.setNodeMarkup($from.before(d), undefined, { ...n.attrs, ...attrs }));
      return;
    }
  }
}

export function applyBulletMain(editor: Editor) {
  toggleList(editor, "bulletList");
}

export function applyNumberMain(editor: Editor) {
  toggleList(editor, "orderedList");
}

export function applyBullet(editor: Editor, style: string | null) {
  if (style === null) {
    if (editor.isActive("bulletList")) toggleList(editor, "bulletList");
    return;
  }
  if (!editor.isActive("bulletList")) toggleList(editor, "bulletList");
  listNodeAttrs(editor, "bulletList", { listStyle: style === "disc" ? null : style, multilevel: null });
  editor.commands.focus();
}

export function applyNumbering(editor: Editor, style: string | null) {
  if (style === null) {
    if (editor.isActive("orderedList")) toggleList(editor, "orderedList");
    return;
  }
  if (!editor.isActive("orderedList")) toggleList(editor, "orderedList");
  listNodeAttrs(editor, "orderedList", { listStyle: style === "decimal" ? null : style, multilevel: null });
  editor.commands.focus();
}

export function applyMultilevel(editor: Editor, id: string) {
  if (id === "bullets") {
    if (!editor.isActive("bulletList")) toggleList(editor, "bulletList");
    listNodeAttrs(editor, "bulletList", { listStyle: null, multilevel: "bullets" });
  } else {
    if (!editor.isActive("orderedList")) toggleList(editor, "orderedList");
    listNodeAttrs(editor, "orderedList", { listStyle: null, multilevel: id });
  }
  editor.commands.focus();
}

function Tile({ selected, onClick, children, title }: { selected?: boolean; onClick: () => void; children: React.ReactNode; title: string }) {
  const close = useMenuClose();
  return (
    <button
      type="button"
      className={`bl-libtile${selected ? " is-selected" : ""}`}
      title={title}
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => {
        onClick();
        close();
      }}
    >
      {children}
    </button>
  );
}

export function BulletLibrary({ editor, current }: { editor: Editor; current: string | null }) {
  return (
    <div className="oh-menu bl-libmenu">
      <MenuHeader>Madde İşareti Kitaplığı</MenuHeader>
      <div className="bl-libgrid">
        <Tile title="Yok" onClick={() => applyBullet(editor, null)} selected={current === null}>
          <span className="bl-lib-none">Yok</span>
        </Tile>
        {BULLET_STYLES.map((b) => (
          <Tile key={b.id} title={b.label} onClick={() => applyBullet(editor, b.id)} selected={current === b.id}>
            <span className="bl-lib-bullet">{b.char}</span>
          </Tile>
        ))}
      </div>
    </div>
  );
}

export function NumberLibrary({ editor, current }: { editor: Editor; current: string | null }) {
  return (
    <div className="oh-menu bl-libmenu">
      <MenuHeader>Numaralandırma Kitaplığı</MenuHeader>
      <div className="bl-libgrid is-numbers">
        <Tile title="Yok" onClick={() => applyNumbering(editor, null)} selected={current === null}>
          <span className="bl-lib-none">Yok</span>
        </Tile>
        {NUMBER_STYLES.map((n) => (
          <Tile key={n.id} title={n.label} onClick={() => applyNumbering(editor, n.id)} selected={current === n.id}>
            <span className="bl-lib-numbers">
              {n.sample.map((s) => (
                <span key={s}>
                  <b>{s}</b>
                  <i />
                </span>
              ))}
            </span>
          </Tile>
        ))}
      </div>
      <MenuSeparator />
      <MenuItem label="Numaralandırmaya Yeniden Başla" onClick={() => editor.chain().focus().updateAttributes("orderedList", { start: 1 }).run()} />
    </div>
  );
}

export function MultilevelLibrary({ editor }: { editor: Editor }) {
  return (
    <div className="oh-menu bl-libmenu">
      <MenuHeader>Liste Kitaplığı</MenuHeader>
      <div className="bl-libgrid is-numbers">
        {MULTILEVEL_STYLES.map((m) => (
          <Tile key={m.id} title={m.label} onClick={() => applyMultilevel(editor, m.id)}>
            <span className="bl-lib-numbers">
              {m.sample.map((s, i) => (
                <span key={s} style={{ paddingLeft: i * 10 }}>
                  <b>{s}</b>
                  <i />
                </span>
              ))}
            </span>
          </Tile>
        ))}
      </div>
      <MenuSeparator />
      <MenuItem label="Liste Düzeyini Artır" shortcut="Tab" onClick={() => editor.chain().focus().sinkListItem("listItem").run()} />
      <MenuItem label="Liste Düzeyini Azalt" shortcut="Shift+Tab" onClick={() => editor.chain().focus().liftListItem("listItem").run()} />
    </div>
  );
}
