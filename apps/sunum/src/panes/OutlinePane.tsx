/**
 * Ana Hat Görünümü: her slaydın başlığı ve gövde metni düzenlenebilir bir ana hat
 * olarak listelenir. Gövde satırlarındaki baştaki sekmeler liste düzeyini belirtir
 * (Tab / Shift+Tab ile değiştirilir).
 */
import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { updateSlideShapes } from "../model/ops";
import { bodyShape, bodyToLines, linesToBody, paraText, titleShape } from "../model/outline";
import type { Shape, Slide, TextBody } from "../model/types";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";

function setShapeText(a: Actions, slideIdx: number, shapeId: string, body: TextBody, key: string) {
  a.store.commit((p) => updateSlideShapes(p, slideIdx, (list: Shape[]) => list.map((x) => (x.id === shapeId && x.kind === "sp" ? { ...x, text: body } : x))), { merge: key });
}

function OutlineSlide({ slide, index, current, actions: a, spellcheck }: { slide: Slide; index: number; current: boolean; actions: Actions; spellcheck: boolean }) {
  const tShape = titleShape(slide);
  const bShape = bodyShape(slide);
  const [title, setTitle] = useState(() => (tShape?.text ? tShape.text.paragraphs.map(paraText).join(" ") : ""));
  const [body, setBody] = useState(() => bodyToLines(bShape?.text));
  const bodyRef = useRef<HTMLTextAreaElement>(null);
  const editing = useRef(false);

  // Dışarıdan (slayt alanında) yapılan değişiklikleri yansıt; yazarken üzerine yazma.
  useEffect(() => {
    if (editing.current) return;
    setTitle(tShape?.text ? tShape.text.paragraphs.map(paraText).join(" ") : "");
    setBody(bodyToLines(bShape?.text));
  }, [tShape?.text, bShape?.text]);

  const focus = () => {
    editing.current = true;
    if (!current) a.goToSlide(index);
    a.store.set({ focus: "notes", editing: null, sel: [] });
  };
  const blur = () => {
    editing.current = false;
    a.store.breakMerge();
  };

  // Satır kaydırmayla birlikte yüksekliği içeriğe uydur.
  useLayoutEffect(() => {
    const el = bodyRef.current;
    if (!el) return;
    el.style.height = "auto";
    el.style.height = `${el.scrollHeight + 2}px`;
  }, [body]);

  const changeLevel = (el: HTMLTextAreaElement, delta: 1 | -1) => {
    const v = el.value;
    const s0 = v.lastIndexOf("\n", el.selectionStart - 1) + 1;
    const e0 = v.indexOf("\n", el.selectionEnd);
    const end = e0 < 0 ? v.length : e0;
    const block = v.slice(s0, end).split("\n");
    const next = block.map((l) => (delta > 0 ? (/^\t{8}/.test(l) ? l : `\t${l}`) : l.replace(/^\t/, ""))).join("\n");
    const nv = v.slice(0, s0) + next + v.slice(end);
    const shift = next.length - (end - s0);
    const selStart = el.selectionStart;
    const selEnd = el.selectionEnd;
    setBody(nv);
    if (bShape) setShapeText(a, index, bShape.id, linesToBody(nv, bShape.text), `outline:b:${slide.id}`);
    const firstShift = delta > 0 ? (/^\t{8}/.test(block[0]) ? 0 : 1) : block[0].startsWith("\t") ? -1 : 0;
    requestAnimationFrame(() => {
      el.selectionStart = Math.max(s0, selStart + firstShift);
      el.selectionEnd = Math.max(el.selectionStart, selEnd + shift);
    });
  };

  return (
    <div className={`outline-slide${current ? " is-current" : ""}${slide.hidden ? " is-hidden" : ""}`} onMouseDown={() => !current && a.goToSlide(index)}>
      <div className="outline-num">
        {index + 1}
        <span className="outline-icon" aria-hidden="true" />
      </div>
      <div className="outline-texts">
        <input
          type="text"
          className="outline-title"
          value={title}
          placeholder={tShape ? "Başlık" : "(Başlık yer tutucusu yok)"}
          disabled={!tShape}
          spellCheck={spellcheck}
          lang="tr"
          onFocus={focus}
          onBlur={blur}
          onChange={(e) => {
            setTitle(e.target.value);
            if (tShape) setShapeText(a, index, tShape.id, linesToBody(e.target.value, tShape.text), `outline:t:${slide.id}`);
          }}
          onKeyDown={(e) => {
            e.stopPropagation();
            if (e.key === "Enter" || e.key === "ArrowDown") {
              e.preventDefault();
              bodyRef.current?.focus();
            }
          }}
        />
        {bShape && (
          <textarea
            ref={bodyRef}
            className="outline-body"
            value={body}
            rows={1}
            placeholder="Metin"
            spellCheck={spellcheck}
            lang="tr"
            onFocus={focus}
            onBlur={blur}
            onChange={(e) => {
              setBody(e.target.value);
              setShapeText(a, index, bShape.id, linesToBody(e.target.value, bShape.text), `outline:b:${slide.id}`);
            }}
            onKeyDown={(e) => {
              if (e.key === "Tab") {
                e.preventDefault();
                changeLevel(e.currentTarget, e.shiftKey ? -1 : 1);
              }
              if (!e.ctrlKey && !e.metaKey && !e.key.startsWith("F")) e.stopPropagation();
            }}
          />
        )}
      </div>
    </div>
  );
}

export function OutlinePane({ state, actions, width }: { state: AppState; actions: Actions; width: number }) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    ref.current?.querySelector(".outline-slide.is-current")?.scrollIntoView({ block: "nearest" });
  }, [state.slideIdx]);
  return (
    <div ref={ref} className="outline-pane oh-scroll" style={{ width }} aria-label="Ana hat">
      {state.pres.slides.map((s, i) => (
        <OutlineSlide key={s.id} slide={s} index={i} current={i === state.slideIdx} actions={actions} spellcheck={state.spellcheck} />
      ))}
    </div>
  );
}
