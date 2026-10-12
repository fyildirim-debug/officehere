// Stiller bölmesi (Giriş › Stiller grubu oku, Alt+Ctrl+Shift+S): tüm stiller, uygulanan stilin vurgusu,
// Değiştir / Sil / Varsayılana Döndür, Yeni Stil ve Önizlemeyi Göster.
import { useState } from "react";
import { ChevronDown16Regular, Dismiss16Regular, TextClearFormatting20Regular, TextEditStyle20Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, MenuProvider, Popover } from "@officehere/ui";
import { useRef } from "react";
import { useBelge } from "../state/context";
import { applyStyle, clearFormatting } from "../editor/commands";
import { STYLE_ORDER, previewStyle, styleIds } from "../model/styles";

export function StylesPane({ onClose, onEdit }: { onClose: () => void; onEdit: (id?: string) => void }) {
  const ctx = useBelge();
  const { styles, editor, fmt } = ctx;
  const [preview, setPreview] = useState(false);
  const current = fmt.charStyle ?? fmt.styleId;
  const custom = ctx.settings.customStyles ?? {};
  const remove = (id: string) =>
    ctx.updateSettings((s) => {
      const next = { ...(s.customStyles ?? {}) };
      delete next[id];
      return { ...s, customStyles: next };
    });
  return (
    <aside className="bl-pane bl-styles-pane" aria-label="Stiller">
      <div className="bl-pane-head">
        <strong>Stiller</strong>
        <button type="button" className="bl-pane-close" aria-label="Kapat" onClick={onClose}>
          <Dismiss16Regular />
        </button>
      </div>
      <div className="bl-styles-list" role="listbox" aria-label="Stil listesi">
        <button type="button" className="bl-styles-item" onClick={() => clearFormatting(editor)}>
          <span className="bl-styles-name">Tümünü Temizle</span>
          <TextClearFormatting20Regular />
        </button>
        {styleIds(styles).map((id) => (
          <StyleRow
            key={id}
            id={id}
            name={styles[id].name}
            kind={styles[id].type}
            active={current === id}
            style={preview ? previewStyle(styles, id) : undefined}
            builtin={(STYLE_ORDER as readonly string[]).includes(id)}
            changed={!!custom[id]}
            onApply={() => applyStyle(editor, styles, id)}
            onEdit={() => onEdit(id)}
            onRemove={() => remove(id)}
          />
        ))}
      </div>
      <div className="bl-styles-foot">
        <label className="bl-check">
          <input type="checkbox" checked={preview} onChange={(e) => setPreview(e.target.checked)} /> Önizlemeyi Göster
        </label>
        <button type="button" className="oh-btn" onClick={() => onEdit(undefined)}>
          <TextEditStyle20Regular /> Yeni Stil
        </button>
      </div>
    </aside>
  );
}

function StyleRow(p: { id: string; name: string; kind: "paragraph" | "character"; active: boolean; style?: React.CSSProperties; builtin: boolean; changed: boolean; onApply: () => void; onEdit: () => void; onRemove: () => void }) {
  const ref = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  return (
    <div className={`bl-styles-item${p.active ? " is-active" : ""}`} role="option" aria-selected={p.active} data-style-id={p.id}>
      <button type="button" className="bl-styles-apply" onMouseDown={(e) => e.preventDefault()} onClick={p.onApply}>
        <span className="bl-styles-name" style={p.style ? { ...p.style, fontSize: undefined } : undefined}>
          {p.name}
        </span>
        <span className="bl-styles-kind" aria-hidden="true">
          {p.kind === "paragraph" ? "¶" : "a"}
        </span>
      </button>
      <button ref={ref} type="button" className="bl-styles-more" aria-label={`${p.name} seçenekleri`} onClick={() => setOpen((o) => !o)}>
        <ChevronDown16Regular />
      </button>
      <Popover open={open} onClose={() => setOpen(false)} anchor={ref} placement="bottom-end" role="menu">
        <MenuProvider onClose={() => setOpen(false)}>
          <Menu width={200}>
            <MenuItem label="Değiştir..." onClick={p.onEdit} />
            {!p.builtin && <MenuItem label={`"${p.name}" Stilini Sil`} onClick={p.onRemove} />}
            {p.builtin && p.changed && <MenuItem label="Varsayılana Döndür" onClick={p.onRemove} />}
          </Menu>
        </MenuProvider>
      </Popover>
    </div>
  );
}
