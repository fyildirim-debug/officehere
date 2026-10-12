/** Seçim Bölmesi: sayfadaki grafik, resim ve şekilleri listeler; seçme, gizleme/gösterme, yeniden adlandırma, sıralama. */
import { useState } from "react";
import { ArrowDown16Regular, ArrowUp16Regular, Dismiss16Regular, Eye16Regular, EyeOff16Regular } from "@fluentui/react-icons";
import { tipProps } from "@officehere/ui";
import type { Controller } from "../controller";
import { objectList, renameObject, setObjectHidden, stepObject } from "../actions/objectClip";

export function SelectionPane({ ctl }: { ctl: Controller }) {
  // Üstteki (öndeki) nesne listenin başında
  const items = objectList(ctl).reverse();
  const [renaming, setRenaming] = useState<string | null>(null);
  const [name, setName] = useState("");
  const close = () => {
    ctl.formatPane = null;
    ctl.emit();
    ctl.focusGrid();
  };
  const all = (hidden: boolean) => {
    for (const o of items) if (o.hidden !== hidden) setObjectHidden(ctl, o.kind, o.id, hidden);
  };
  const sel = items.find((o) => o.id === ctl.selObj?.id);
  const commit = () => {
    const o = items.find((x) => x.id === renaming);
    if (o && name.trim() && name.trim() !== o.name) renameObject(ctl, o.kind, o.id, name);
    setRenaming(null);
  };
  return (
    <aside className="tb-pane tb-sel-pane" aria-label="Seçim">
      <div className="tb-pane-head">
        <span>Seçim</span>
        <button type="button" className="tb-pane-close" aria-label="Kapat" onClick={close}>
          <Dismiss16Regular />
        </button>
      </div>
      <div className="tb-sel-actions">
        <button type="button" className="oh-btn" onClick={() => all(false)} disabled={!items.length}>
          Tümünü Göster
        </button>
        <button type="button" className="oh-btn" onClick={() => all(true)} disabled={!items.length}>
          Tümünü Gizle
        </button>
        <span className="tb-sel-order">
          <button type="button" className="tb-pane-close" aria-label="Öne Getir" disabled={!sel} onClick={() => sel && stepObject(ctl, sel.kind, sel.id, true)} {...tipProps("Öne Getir")}>
            <ArrowUp16Regular />
          </button>
          <button type="button" className="tb-pane-close" aria-label="Arkaya Gönder" disabled={!sel} onClick={() => sel && stepObject(ctl, sel.kind, sel.id, false)} {...tipProps("Arkaya Gönder")}>
            <ArrowDown16Regular />
          </button>
        </span>
      </div>
      <div className="tb-sel-list" role="listbox" aria-label="Nesneler">
        {!items.length && <div className="tb-sel-empty">Bu sayfada nesne yok.</div>}
        {items.map((o) => (
          <div
            key={o.id}
            role="option"
            aria-selected={ctl.selObj?.id === o.id}
            className={`tb-sel-item${ctl.selObj?.id === o.id ? " is-active" : ""}${o.hidden ? " is-hidden" : ""}`}
            onClick={() => {
              if (o.hidden) return;
              ctl.selObj = { kind: o.kind, id: o.id };
              ctl.emit();
            }}
            onDoubleClick={() => {
              setRenaming(o.id);
              setName(o.name);
            }}
          >
            {renaming === o.id ? (
              <input
                className="tb-sel-rename"
                aria-label="Nesne adı"
                autoFocus
                value={name}
                onChange={(e) => setName(e.target.value)}
                onFocus={(e) => e.currentTarget.select()}
                onBlur={commit}
                onClick={(e) => e.stopPropagation()}
                onKeyDown={(e) => {
                  e.stopPropagation();
                  if (e.key === "Enter") {
                    e.preventDefault();
                    commit();
                  } else if (e.key === "Escape") setRenaming(null);
                }}
              />
            ) : (
              <span className="tb-sel-name">{o.name}</span>
            )}
            <button type="button" className="tb-pane-close" aria-label={o.hidden ? "Göster" : "Gizle"} onClick={(e) => (e.stopPropagation(), setObjectHidden(ctl, o.kind, o.id, !o.hidden))}>
              {o.hidden ? <EyeOff16Regular /> : <Eye16Regular />}
            </button>
          </div>
        ))}
      </div>
    </aside>
  );
}
