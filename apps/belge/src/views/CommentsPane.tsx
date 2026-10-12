// Sağ kenardaki açıklama balonları.
import { useEffect, useLayoutEffect, useRef, useState } from "react";
import type { Editor } from "@tiptap/core";
import { TextSelection } from "@tiptap/pm/state";
import { Delete16Regular, Checkmark16Regular, Dismiss16Regular } from "@fluentui/react-icons";
import type { CommentData } from "../model/document";
import { visibleCoords, type RevisionBalloon } from "./revisions";

export interface CommentAnchor {
  id: string;
  from: number;
  to: number;
}

/** Belgedeki açıklama işaretlerinin aralıkları (belge sırasıyla). */
export function commentAnchors(editor: Editor): CommentAnchor[] {
  const map = new Map<string, CommentAnchor>();
  editor.state.doc.descendants((node, pos) => {
    for (const m of node.marks) {
      if (m.type.name !== "comment") continue;
      const id = m.attrs.id as string;
      const cur = map.get(id);
      if (cur) cur.to = pos + node.nodeSize;
      else map.set(id, { id, from: pos, to: pos + node.nodeSize });
    }
  });
  return [...map.values()].sort((a, b) => a.from - b.from);
}

interface Props {
  editor: Editor;
  comments: Record<string, CommentData>;
  anchors: CommentAnchor[];
  /** Balonların konumlandırılacağı kapsayıcı (sayfalar). */
  container: HTMLElement | null;
  scale: number;
  left: number;
  activeId: string | null;
  onActivate: (id: string | null) => void;
  onChange: (id: string, text: string) => void;
  onDelete: (id: string) => void;
  onResolve: (id: string) => void;
  version: number;
  /** Düzeltme balonları (silinen metin, biçim değişiklikleri). */
  revisions?: RevisionBalloon[];
  onRevision?: (r: RevisionBalloon, accept: boolean) => void;
}

function fmtDate(iso: string) {
  try {
    return new Date(iso).toLocaleString("tr-TR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
  } catch {
    return iso;
  }
}

export function CommentsPane({ editor, comments, anchors, container, scale, left, activeId, onActivate, onChange, onDelete, onResolve, version, revisions = [], onRevision }: Props) {
  const [tops, setTops] = useState<Record<string, number>>({});
  const refs = useRef<Record<string, HTMLDivElement | null>>({});

  useLayoutEffect(() => {
    if (!container) return;
    const base = container.getBoundingClientRect().top;
    const want: { id: string; y: number }[] = [];
    for (const a of [...anchors, ...revisions].sort((x, y) => x.from - y.from)) {
      const c = visibleCoords(editor.view, Math.min(a.from + 1, a.to));
      if (c) want.push({ id: a.id, y: (c.top - base) / scale });
    }
    // Çakışmayı önlemek için aşağı it.
    const out: Record<string, number> = {};
    let floor = -Infinity;
    for (const w of want) {
      const h = (refs.current[w.id]?.offsetHeight ?? 60) + 8;
      const y = Math.max(w.y - 6, floor);
      out[w.id] = y;
      floor = y + h;
    }
    setTops((prev) => (JSON.stringify(prev) === JSON.stringify(out) ? prev : out));
  }, [anchors, container, scale, editor, version, comments, revisions]);

  return (
    <div className="bl-comments" style={{ left }}>
      {revisions.map((r) => (
        <div
          key={r.id}
          ref={(el) => {
            refs.current[r.id] = el;
          }}
          className={`bl-balloon is-revision is-${r.kind}`}
          style={{ top: tops[r.id] ?? 0 }}
          onMouseDown={(e) => {
            if ((e.target as HTMLElement).closest("button")) return;
            e.preventDefault();
            editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, r.from, Math.min(r.to, editor.state.doc.content.size))).scrollIntoView());
            editor.view.focus();
          }}
        >
          <div className="bl-balloon-head">
            <span className="bl-balloon-author">{r.author || "Yazar"}</span>
            <span className="bl-balloon-date">{fmtDate(r.date)}</span>
            {onRevision && (
              <>
                <button type="button" className="bl-balloon-btn" title="Kabul Et" onClick={() => onRevision(r, true)}>
                  <Checkmark16Regular />
                </button>
                <button type="button" className="bl-balloon-btn" title="Reddet" onClick={() => onRevision(r, false)}>
                  <Dismiss16Regular />
                </button>
              </>
            )}
          </div>
          <div className="bl-balloon-rev">
            <b>{r.label}</b>
            {r.text && <span className="bl-balloon-deleted">{r.text}</span>}
          </div>
        </div>
      ))}
      {anchors.map((a) => {
        const c = comments[a.id];
        if (!c) return null;
        const active = activeId === a.id;
        return (
          <div
            key={a.id}
            ref={(el) => {
              refs.current[a.id] = el;
            }}
            className={`bl-balloon${active ? " is-active" : ""}${c.resolved ? " is-resolved" : ""}`}
            style={{ top: tops[a.id] ?? 0 }}
            onMouseDown={(e) => {
              if ((e.target as HTMLElement).closest("textarea,button")) return;
              e.preventDefault();
              onActivate(a.id);
              const tr = editor.state.tr.setSelection(TextSelection.create(editor.state.doc, a.from, a.to));
              editor.view.dispatch(tr);
            }}
          >
            <div className="bl-balloon-head">
              <span className="bl-avatar">{c.initials}</span>
              <span className="bl-balloon-author">{c.author}</span>
              <span className="bl-balloon-date">{fmtDate(c.date)}</span>
              <button type="button" className="bl-balloon-btn" title="Çözümle" onClick={() => onResolve(a.id)}>
                <Checkmark16Regular />
              </button>
              <button type="button" className="bl-balloon-btn" title="Açıklamayı Sil" onClick={() => onDelete(a.id)}>
                <Delete16Regular />
              </button>
            </div>
            <CommentText text={c.text} autoFocus={active && !c.text} onChange={(t) => onChange(a.id, t)} onFocus={() => onActivate(a.id)} />
          </div>
        );
      })}
    </div>
  );
}

function CommentText({ text, autoFocus, onChange, onFocus }: { text: string; autoFocus: boolean; onChange: (t: string) => void; onFocus: () => void }) {
  const ref = useRef<HTMLTextAreaElement>(null);
  useEffect(() => {
    if (autoFocus) ref.current?.focus();
  }, [autoFocus]);
  useLayoutEffect(() => {
    const el = ref.current;
    if (el) {
      el.style.height = "auto";
      el.style.height = `${el.scrollHeight}px`;
    }
  }, [text]);
  return (
    <textarea
      ref={ref}
      className="bl-balloon-text"
      value={text}
      placeholder="Açıklama ekleyin…"
      rows={1}
      onFocus={onFocus}
      onChange={(e) => onChange(e.target.value)}
    />
  );
}
