/** Animasyon Bölmesi, Seçim Bölmesi ve Açıklamalar bölmesi. */
import { useState } from "react";
import { ArrowDown20Regular, ArrowUp20Regular, Delete20Regular, Eye20Regular, EyeOff20Regular, Play20Regular } from "@fluentui/react-icons";
import { getUserName } from "@officehere/ui";
import { formatDateTime } from "@officehere/core";
import { findShape, mapShapes, reorder, updateSlide } from "../model/ops";
import type { Animation, Shape } from "../model/types";
import { animDef, TRIGGER_LABELS } from "../show/catalog";
import { AnimStar } from "../ribbon/AnimationsTab";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";
import { PaneShell } from "./paneControls";

export function AnimationPane({ state, actions: a, onClose, onPlay }: { state: AppState; actions: Actions; onClose: () => void; onPlay: (fromId?: string) => void }) {
  const slide = state.pres.slides[state.slideIdx];
  const anims = slide?.animations ?? [];
  const selected = anims.filter((x) => state.sel.includes(x.shapeId)).map((x) => x.id);
  let click = 0;
  const move = (id: string, dir: -1 | 1) =>
    a.setAnimations((list) => {
      const i = list.findIndex((x) => x.id === id);
      const j = i + dir;
      if (i < 0 || j < 0 || j >= list.length) return list;
      const out = list.slice();
      [out[i], out[j]] = [out[j], out[i]];
      return out;
    });
  return (
    <PaneShell title="Animasyon Bölmesi" onClose={onClose}>
      <div className="pane-toolbar">
        <button type="button" className="oh-btn" disabled={!anims.length} onClick={() => onPlay(selected[0])}>
          <Play20Regular style={{ width: 14, verticalAlign: -3 }} /> {selected.length ? "Seçilenden Yürüt" : "Tümünü Yürüt"}
        </button>
        <span style={{ flex: 1 }} />
        <button type="button" className="icon-btn" title="Yukarı taşı" disabled={selected.length !== 1} onClick={() => move(selected[0], -1)}>
          <ArrowUp20Regular />
        </button>
        <button type="button" className="icon-btn" title="Aşağı taşı" disabled={selected.length !== 1} onClick={() => move(selected[0], 1)}>
          <ArrowDown20Regular />
        </button>
      </div>
      {!anims.length && <p style={{ color: "var(--oh-text-muted)" }}>Bu slaytta animasyon yok. Bir nesne seçip Animasyonlar sekmesinden efekt uygulayın.</p>}
      <div className="list-pane">
        {anims.map((an: Animation) => {
          if (an.trigger === "click") click += 1;
          const sh = slide ? findShape(slide.shapes, an.shapeId) : undefined;
          const def = animDef(an.effect);
          const sel = state.sel.includes(an.shapeId);
          return (
            <div key={an.id} className={`list-row${sel ? " is-selected" : ""}`} onClick={() => a.select([an.shapeId])} title={`${def?.label ?? an.effect} — ${TRIGGER_LABELS[an.trigger]}`}>
              <span style={{ width: 16, textAlign: "right", color: "var(--oh-text-muted)" }}>{an.trigger === "click" ? click : ""}</span>
              <AnimStar cls={an.cls} size={16} />
              <span className="grow">{sh?.name || "(silinmiş)"}</span>
              <select
                value={an.trigger}
                onClick={(e) => e.stopPropagation()}
                onChange={(e) => a.setAnimations((list) => list.map((x) => (x.id === an.id ? { ...x, trigger: e.target.value as Animation["trigger"] } : x)))}
                title="Başlangıç"
                style={{ height: 22, fontSize: 11 }}
              >
                {Object.entries(TRIGGER_LABELS).map(([k, l]) => (
                  <option key={k} value={k}>
                    {l}
                  </option>
                ))}
              </select>
              <button type="button" className="icon-btn" title="Kaldır" onClick={(e) => (e.stopPropagation(), a.setAnimations((list) => list.filter((x) => x.id !== an.id)))}>
                <Delete20Regular />
              </button>
            </div>
          );
        })}
      </div>
    </PaneShell>
  );
}

export function SelectionPane({ state, actions: a, onClose }: { state: AppState; actions: Actions; onClose: () => void }) {
  const slide = state.pres.slides[state.slideIdx];
  const [editing, setEditing] = useState<string | null>(null);
  const setHidden = (ids: Set<string> | ((s: Shape) => boolean), v: boolean) => a.store.commit((p) => updateSlide(p, state.slideIdx, (sl) => ({ ...sl, shapes: mapShapes(sl.shapes, ids, (s) => ({ ...s, hidden: v || undefined })) })));
  const rename = (id: string, name: string) => a.store.commit((p) => updateSlide(p, state.slideIdx, (sl) => ({ ...sl, shapes: mapShapes(sl.shapes, new Set([id]), (s) => ({ ...s, name })) })));
  const list = slide ? [...slide.shapes].reverse() : [];
  const row = (s: Shape, depth: number) => (
    <div key={s.id}>
      <div className={`list-row${state.sel.includes(s.id) ? " is-selected" : ""}`} style={{ paddingLeft: 6 + depth * 14 }} onClick={(e) => a.store.set({ sel: e.ctrlKey ? [...state.sel, s.id] : [s.id], editing: null })} onDoubleClick={() => setEditing(s.id)}>
        {editing === s.id ? (
          <input
            autoFocus
            defaultValue={s.name}
            className="grow"
            onBlur={(e) => {
              rename(s.id, e.target.value);
              setEditing(null);
            }}
            onKeyDown={(e) => {
              e.stopPropagation();
              if (e.key === "Enter") (e.target as HTMLInputElement).blur();
              if (e.key === "Escape") setEditing(null);
            }}
          />
        ) : (
          <span className="grow" style={{ opacity: s.hidden ? 0.5 : 1 }}>
            {s.name || "(adsız)"}
          </span>
        )}
        <button type="button" className="icon-btn" title={s.hidden ? "Göster" : "Gizle"} onClick={(e) => (e.stopPropagation(), setHidden(new Set([s.id]), !s.hidden))}>
          {s.hidden ? <EyeOff20Regular /> : <Eye20Regular />}
        </button>
      </div>
      {s.kind === "group" && [...s.children].reverse().map((c) => row(c, depth + 1))}
    </div>
  );
  return (
    <PaneShell title="Seçim" onClose={onClose}>
      <div className="pane-toolbar">
        <button type="button" className="oh-btn" onClick={() => setHidden(() => true, false)}>
          Tümünü Göster
        </button>
        <button type="button" className="oh-btn" onClick={() => setHidden((s) => slide?.shapes.some((x) => x.id === s.id) ?? false, true)}>
          Tümünü Gizle
        </button>
        <span style={{ flex: 1 }} />
        <button type="button" className="icon-btn" title="Öne Getir" disabled={!state.sel.length} onClick={() => a.store.commit((p) => updateSlide(p, state.slideIdx, (sl) => ({ ...sl, shapes: reorder(sl.shapes, state.sel, "forward") })))}>
          <ArrowUp20Regular />
        </button>
        <button type="button" className="icon-btn" title="Arkaya Gönder" disabled={!state.sel.length} onClick={() => a.store.commit((p) => updateSlide(p, state.slideIdx, (sl) => ({ ...sl, shapes: reorder(sl.shapes, state.sel, "backward") })))}>
          <ArrowDown20Regular />
        </button>
      </div>
      <div className="list-pane">{list.map((s) => row(s, 0))}</div>
    </PaneShell>
  );
}

export function CommentsPane({ state, actions: a, onClose }: { state: AppState; actions: Actions; onClose: () => void }) {
  const slide = state.pres.slides[state.slideIdx];
  const comments = slide?.comments ?? [];
  const update = (id: string, text: string | null) =>
    a.store.commit(
      (p) => updateSlide(p, state.slideIdx, (sl) => ({ ...sl, comments: (sl.comments ?? []).flatMap((c) => (c.id !== id ? [c] : text === null ? [] : [{ ...c, text }])) })),
      { merge: `comment:${id}` },
    );
  return (
    <PaneShell title="Açıklamalar" onClose={onClose}>
      <div className="pane-toolbar">
        <button
          type="button"
          className="oh-btn"
          disabled={!slide}
          onClick={() => a.addComment(getUserName())}
        >
          Yeni
        </button>
      </div>
      {!comments.length && <p style={{ color: "var(--oh-text-muted)" }}>Bu slaytta açıklama yok.</p>}
      {comments.map((c) => (
        <div key={c.id} className="fp-section" style={{ padding: 8 }}>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
            <strong>{c.author}</strong>
            <span style={{ color: "var(--oh-text-muted)", fontSize: 11 }}>{formatDateTime(c.date)}</span>
          </div>
          <textarea
            value={c.text}
            placeholder="Açıklama yazın..."
            rows={3}
            style={{ width: "100%", resize: "vertical", border: "1px solid var(--oh-border)", borderRadius: 4, padding: 6, background: "var(--oh-surface)" }}
            onChange={(e) => update(c.id, e.target.value)}
            onKeyDown={(e) => e.stopPropagation()}
          />
          <button type="button" className="icon-btn" title="Sil" onClick={() => update(c.id, null)}>
            <Delete20Regular /> Sil
          </button>
        </div>
      ))}
    </PaneShell>
  );
}
