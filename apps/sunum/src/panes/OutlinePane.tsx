/**
 * Ana Hat Görünümü: her slaydın başlığı ve gövde metni düzenlenebilir bir ana hat
 * olarak listelenir. Gövde satırlarındaki baştaki sekmeler liste düzeyini belirtir
 * (Tab / Shift+Tab ile değiştirilir).
 */
import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { updateSlideShapes } from "../model/ops";
import { bodyShape, bodyToLines, demoteSlide, linesToBody, paraText, promoteLine, titleShape } from "../model/outline";
import { newSlide as makeSlide } from "../model/presentation";
import type { Shape, Slide, TextBody } from "../model/types";
import type { Actions } from "../state/actions";
import type { AppState } from "../state/store";

function setShapeText(a: Actions, slideIdx: number, shapeId: string, body: TextBody, key: string) {
  a.store.commit((p) => updateSlideShapes(p, slideIdx, (list: Shape[]) => list.map((x) => (x.id === shapeId && x.kind === "sp" ? { ...x, text: body } : x))), { merge: key });
}

/** Yapı değişikliğinden (yeni slayt, indirme/yükseltme) sonra odaklanacak alan. */
let pendingFocus: { slideId: string; field: "title" | "body"; caret?: number } | null = null;

function OutlineSlide({ slide, index, current, actions: a, spellcheck }: { slide: Slide; index: number; current: boolean; actions: Actions; spellcheck: boolean }) {
  const tShape = titleShape(slide);
  const bShape = bodyShape(slide);
  const [title, setTitle] = useState(() => (tShape?.text ? tShape.text.paragraphs.map(paraText).join(" ") : ""));
  const [body, setBody] = useState(() => bodyToLines(bShape?.text));
  const bodyRef = useRef<HTMLTextAreaElement>(null);
  const titleRef = useRef<HTMLInputElement>(null);
  const editing = useRef(false);
  /** Alanları modelden yeniden doldurmak için (odak çıkınca, geri al/yinele sonrası). */
  const [syncTick, setSyncTick] = useState(0);


  /** Geri al/yinele: uygulama geçmişi kullanılır, alanlar modelden yeniden doldurulur. */
  const history = (e: React.KeyboardEvent) => {
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key.toLowerCase();
    if (!mod || e.altKey || (k !== "z" && k !== "y")) return false;
    e.preventDefault();
    e.stopPropagation();
    editing.current = false;
    a.store.breakMerge();
    if (k === "y" || e.shiftKey) a.store.redo();
    else a.store.undo();
    setSyncTick((t) => t + 1);
    return true;
  };

  /** Bu slayttan sonra yeni slayt; başlığına odaklanılır. */
  const newSlideAfter = () => {
    a.goToSlide(index);
    a.newSlide();
    const ns = a.s.pres.slides[index + 1];
    if (ns) pendingFocus = { slideId: ns.id, field: "title" };
  };

  // Dışarıdan (slayt alanında) yapılan değişiklikleri yansıt; yazarken üzerine yazma.
  useEffect(() => {
    if (editing.current) return;
    setTitle(tShape?.text ? tShape.text.paragraphs.map(paraText).join(" ") : "");
    setBody(bodyToLines(bShape?.text));
  }, [tShape?.text, bShape?.text, syncTick]);

  // Yapı değişikliğinden sonra bekleyen odak: alanlar önce modelden doldurulur.
  useEffect(() => {
    const f = pendingFocus;
    if (!f || f.slideId !== slide.id) return;
    pendingFocus = null;
    const el = f.field === "title" ? titleRef.current : bodyRef.current;
    if (!el) return;
    const value = f.field === "title" ? (tShape?.text ? tShape.text.paragraphs.map(paraText).join(" ") : "") : bodyToLines(bShape?.text);
    if (f.field === "title") setTitle(value);
    else setBody(value);
    el.value = value;
    el.focus();
    const c = f.caret ?? value.length;
    el.setSelectionRange(c, c);
  });

  const focus = () => {
    editing.current = true;
    if (!current) a.goToSlide(index);
    a.store.set({ focus: "notes", editing: null, sel: [] });
  };
  const blur = () => {
    editing.current = false;
    a.store.breakMerge();
    // Odak varken yapılan yapı değişiklikleri (ör. satırın yeni slayda taşınması) alana yansır.
    setSyncTick((t) => t + 1);
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
    // İmleç hemen konur (yeniden çizimden önce): ardından yazılan karakterler doğru yere gider.
    el.value = nv;
    el.selectionStart = Math.max(s0, selStart + firstShift);
    el.selectionEnd = Math.max(el.selectionStart, selEnd + shift);
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
          ref={titleRef}
          onChange={(e) => {
            editing.current = true;
            setTitle(e.target.value);
            if (tShape) setShapeText(a, index, tShape.id, linesToBody(e.target.value, tShape.text), `outline:t:${slide.id}`);
          }}
          onKeyDown={(e) => {
            if (history(e)) return;
            e.stopPropagation();
            if (e.key === "Enter" && !e.shiftKey && !e.altKey) {
              // Başlıkta Enter yeni slayt ekler; Ctrl+Enter gövdeye geçer.
              e.preventDefault();
              if (e.ctrlKey || e.metaKey) bodyRef.current?.focus();
              else newSlideAfter();
            } else if (e.key === "ArrowDown") {
              e.preventDefault();
              bodyRef.current?.focus();
            } else if (e.key === "Tab" && !e.shiftKey) {
              // Başlıkta Tab: slayt önceki slaydın gövdesine indirilir.
              e.preventDefault();
              const prevId = a.s.pres.slides[index - 1]?.id;
              const np = demoteSlide(a.s.pres, index);
              if (!np || !prevId) return;
              a.store.commit(() => np, { slideIdx: index - 1, sel: [], editing: null });
              pendingFocus = { slideId: prevId, field: "body" };
            } else if (e.key === "Tab") e.preventDefault();
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
              editing.current = true;
              setBody(e.target.value);
              setShapeText(a, index, bShape.id, linesToBody(e.target.value, bShape.text), `outline:b:${slide.id}`);
            }}
            onKeyDown={(e) => {
              if (history(e)) return;
              const el = e.currentTarget;
              if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
                // Gövdede Ctrl+Enter yeni slayt ekler.
                e.preventDefault();
                e.stopPropagation();
                newSlideAfter();
                return;
              }
              if (e.key === "Enter" && !e.shiftKey && !e.altKey) {
                // Yeni satır geçerli satırın düzeyinde başlar.
                e.preventDefault();
                e.stopPropagation();
                const v = el.value;
                const lineStart = v.lastIndexOf("\n", el.selectionStart - 1) + 1;
                const tabs = /^\t*/.exec(v.slice(lineStart))![0];
                const ins = `\n${tabs}`;
                const nv = v.slice(0, el.selectionStart) + ins + v.slice(el.selectionEnd);
                const caretAt = el.selectionStart + ins.length;
                editing.current = true;
                // Değer ve imleç hemen yazılır (React yeniden çizince imleç yerinde kalır; hızlı yazımda kayma olmaz).
                el.value = nv;
                el.setSelectionRange(caretAt, caretAt);
                setBody(nv);
                setShapeText(a, index, bShape.id, linesToBody(nv, bShape.text), `outline:b:${slide.id}`);
                return;
              }
              if (e.key === "Tab") {
                e.preventDefault();
                const line = el.value.slice(0, el.selectionStart).split("\n").length - 1;
                const top = !el.value.split("\n")[line]?.startsWith("\t");
                if (e.shiftKey && top && el.selectionStart === el.selectionEnd) {
                  // İlk düzeydeki satırda Shift+Tab: satır yeni slaydın başlığı olur.
                  const np = promoteLine(a.s.pres, index, line, () => makeSlide(a.s.pres, a.s.pres.slides[index].layoutId));
                  if (np) {
                    a.store.commit(() => np, { slideIdx: index + 1, sel: [], editing: null });
                    pendingFocus = { slideId: np.slides[index + 1].id, field: "title" };
                  }
                } else changeLevel(el, e.shiftKey ? -1 : 1);
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
