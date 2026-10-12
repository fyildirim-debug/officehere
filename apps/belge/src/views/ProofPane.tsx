// Düzenleyici bölmesi (Gözden Geçir › Yazım ve Dil Bilgisi, F7): belge puanı yerine sorun sayıları,
// iyileştirme önerileri listesi (tıklayınca seçilir, Düzelt / Yoksay) ve yazım denetimi seçeneği.
import { useMemo, useState } from "react";
import type { Editor } from "@tiptap/core";
import { TextSelection } from "@tiptap/pm/state";
import { Dismiss16Regular, CheckmarkCircle20Regular } from "@fluentui/react-icons";
import { findIssues, type ProofIssue } from "../editor/proofing";
import { docStats } from "../editor/commands";

interface Props {
  editor: Editor;
  version: number;
  spellcheck: boolean;
  setSpellcheck: (v: boolean) => void;
  onClose: () => void;
}

export function ProofPane({ editor, version, spellcheck, setSpellcheck, onClose }: Props) {
  const [ignored, setIgnored] = useState<Set<string>>(new Set());
  const key = (i: ProofIssue) => `${i.kind}:${i.from}:${i.text}`;
  const issues = useMemo(() => findIssues(editor.state.doc).filter((i) => !ignored.has(key(i))), [editor, version, ignored]); // eslint-disable-line react-hooks/exhaustive-deps
  const stats = useMemo(() => docStats(editor.state.doc), [editor, version]); // eslint-disable-line react-hooks/exhaustive-deps
  const counts = issues.reduce<Record<string, number>>((m, i) => ((m[i.title] = (m[i.title] ?? 0) + 1), m), {});

  const select = (i: ProofIssue) => {
    const { doc } = editor.state;
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(doc, Math.min(i.from, doc.content.size), Math.min(i.to, doc.content.size))).scrollIntoView());
    editor.view.focus();
  };
  const fix = (i: ProofIssue) => {
    if (i.fix === null) return;
    editor.view.dispatch(editor.state.tr.insertText(i.fix, i.from, i.to).scrollIntoView());
  };

  return (
    <aside className="bl-pane bl-proof" aria-label="Düzenleyici">
      <div className="bl-pane-head">
        <strong>Düzenleyici</strong>
        <button type="button" className="bl-pane-close" aria-label="Kapat" onClick={onClose}>
          <Dismiss16Regular />
        </button>
      </div>
      <div className="bl-proof-summary">
        <div className="bl-proof-score">{issues.length ? `${issues.length} öneri` : "Öneri yok"}</div>
        <div className="bl-proof-sub">
          {stats.words.toLocaleString("tr-TR")} sözcük · {stats.paragraphs.toLocaleString("tr-TR")} paragraf
        </div>
      </div>
      <label className="bl-check bl-proof-spell">
        <input type="checkbox" checked={spellcheck} onChange={(e) => setSpellcheck(e.target.checked)} />
        Yazım hatalarının altını çiz (tarayıcı yazım denetleyicisi)
      </label>
      {Object.keys(counts).length > 0 && (
        <div className="bl-proof-cats">
          {Object.entries(counts).map(([t, n]) => (
            <span key={t}>
              {t}: <b>{n}</b>
            </span>
          ))}
        </div>
      )}
      <div className="bl-proof-list" role="list">
        {issues.length === 0 && (
          <div className="bl-proof-empty">
            <CheckmarkCircle20Regular /> Belgede iyileştirme önerisi bulunamadı.
          </div>
        )}
        {issues.slice(0, 100).map((i) => (
          <div key={key(i)} role="listitem" className="bl-proof-item" onClick={() => select(i)}>
            <div className="bl-proof-title">{i.title}</div>
            <div className="bl-proof-text">
              <span className="bl-proof-bad">{i.text.replace(/ /g, "·")}</span>
              {i.fix !== null && (
                <>
                  {" → "}
                  <span className="bl-proof-good">{i.fix === " " ? "·" : i.fix}</span>
                </>
              )}
            </div>
            <div className="bl-proof-actions">
              {i.fix !== null && (
                <button
                  type="button"
                  className="oh-btn"
                  onClick={(e) => {
                    e.stopPropagation();
                    fix(i);
                  }}
                >
                  Düzelt
                </button>
              )}
              <button
                type="button"
                className="oh-btn"
                onClick={(e) => {
                  e.stopPropagation();
                  setIgnored((s) => new Set(s).add(key(i)));
                }}
              >
                Yoksay
              </button>
            </div>
          </div>
        ))}
      </div>
    </aside>
  );
}
