/**
 * Şekil/hücre metin düzenleyicisi. Model kaynaktır: tarayıcının girdi
 * olayları (beforeinput) engellenip modele uygulanır, DOM modelden yeniden
 * çizilir ve imleç geri konur. IME bileşimi tarayıcıya bırakılır.
 */
import { useEffect, useLayoutEffect, useRef } from "react";
import {
  applyParaProps,
  applyRunProps,
  caret,
  cmpPos,
  deleteRange,
  insertParagraphs,
  insertText,
  isCollapsed,
  orderSel,
  paraLen,
  propsAt,
  removeBackward,
  removeForward,
  runPropsInRange,
  selectAll,
  sliceBody,
  splitParagraph,
  textInRange,
  transformText,
  wordBoundary,
  wordRangeAt,
  type Pos,
  type Sel,
} from "../model/text";
import { effectiveRun } from "../model/textStyle";
import type { Paragraph, RunProps, TextBody } from "../model/types";
import { paragraphsHtml, type TextRenderCtx } from "../render/textHtml";
import { applySel, domSelToSel, domToPos, posFromPoint } from "./domMap";

export interface TextEditorHandle {
  applyRun(patch: (p: RunProps, para: Paragraph) => RunProps): void;
  applyPara(patch: (p: Paragraph) => Paragraph, levelDelta?: number): void;
  /** Seçimdeki çözülmüş (kalıtımlı) biçimler. */
  runProps(): RunProps[];
  /** Seçimdeki paragraflar. */
  paragraphs(): Paragraph[];
  transform(fn: (s: string, c: { paraStart: boolean }) => string): void;
  insert(text: string): void;
  /** Biçimli çalıştırma ekler (alan: slayt numarası vb.). */
  insertRun(text: string, props: RunProps): void;
  /** Biçimli paragrafları yapıştırır. */
  insertRich(paras: Paragraph[]): void;
  selectAll(): void;
  selection(): Sel;
  setSelection(sel: Sel): void;
  selectedText(): string;
  focus(): void;
  /** Ekran noktasına en yakın konuma imleç koyar (metin bloğu dışındaki tıklamalar için). */
  placeCaret(x: number, y: number): void;
}

export interface TextEditorProps {
  body: TextBody;
  ctx: TextRenderCtx;
  initial?: { selectAll?: boolean; replaceWith?: string; point?: { x: number; y: number }; range?: [Pos, Pos] };
  onChange: (body: TextBody, merge?: string) => void;
  onExit: () => void;
  register: (h: TextEditorHandle | null) => void;
  onSelectionChange: () => void;
  onTab?: (shift: boolean) => boolean;
  onUndo: () => void;
  onRedo: () => void;
  onCopyRich: (paras: Paragraph[]) => void;
  getRichClipboard: () => Paragraph[] | null;
  spellcheck: boolean;
  /** Yazım denetimi dili. */
  lang?: string;
  /** En yüksek liste düzeyi (yer tutucuda 4 = 5 düzey, diğerlerinde 8). */
  maxLevel?: number;
}

const RICH_MIME = "application/x-officehere-sunum-text";

export function TextEditor(props: TextEditorProps) {
  const ref = useRef<HTMLDivElement>(null);
  const bodyRef = useRef(props.body);
  const selRef = useRef<Sel>(caret(0, 0));
  const htmlRef = useRef("");
  const pendingRef = useRef<RunProps | null>(null);
  const composing = useRef(false);
  const propsRef = useRef(props);
  propsRef.current = props;

  // Dışarıdan gelen gövde (geri al vb.) yerel kopyayı günceller.
  if (props.body !== bodyRef.current) bodyRef.current = props.body;

  const render = (force = false) => {
    const el = ref.current;
    if (!el) return;
    const html = paragraphsHtml(bodyRef.current, propsRef.current.ctx);
    if (force || html !== htmlRef.current) {
      el.innerHTML = html;
      htmlRef.current = html;
    }
  };

  const clampSel = (s: Sel): Sel => {
    const b = bodyRef.current;
    const fix = (p: Pos): Pos => {
      const pi = Math.max(0, Math.min(b.paragraphs.length - 1, p.p));
      return { p: pi, o: Math.max(0, Math.min(paraLen(b.paragraphs[pi]), p.o)) };
    };
    return { anchor: fix(s.anchor), focus: fix(s.focus) };
  };

  const restoreSel = () => {
    const el = ref.current;
    if (!el || document.activeElement !== el) return;
    selRef.current = clampSel(selRef.current);
    applySel(el, selRef.current);
  };

  const commit = (body: TextBody, sel: Sel, merge?: string) => {
    bodyRef.current = body;
    selRef.current = sel;
    pendingRef.current = null;
    render();
    restoreSel();
    propsRef.current.onChange(body, merge);
    propsRef.current.onSelectionChange();
  };

  useLayoutEffect(() => {
    render();
    restoreSel();
  });

  // İlk açılış: odak ve imleç.
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    render(true);
    el.focus({ preventScroll: true });
    const init = propsRef.current.initial;
    const b = bodyRef.current;
    if (init?.replaceWith !== undefined) {
      const all = selectAll(b);
      const props0 = propsAt(b, { p: 0, o: 0 });
      const cleared = deleteRange(b, all.anchor, all.focus);
      const r = insertText(cleared, caret(0, 0), init.replaceWith, props0);
      commit(r.body, caret(r.caret.p, r.caret.o), `text`);
      return;
    }
    if (init?.range) selRef.current = clampSel({ anchor: init.range[0], focus: init.range[1] });
    else if (init?.selectAll) selRef.current = selectAll(b);
    else if (init?.point) {
      const p = posFromPoint(el, init.point.x, init.point.y);
      selRef.current = p ? { anchor: p, focus: p } : selectAll(b);
      if (!p) {
        const last = b.paragraphs.length - 1;
        selRef.current = caret(last, paraLen(b.paragraphs[last]));
      }
    } else {
      const last = b.paragraphs.length - 1;
      selRef.current = caret(last, paraLen(b.paragraphs[last]));
    }
    applySel(el, selRef.current);
    propsRef.current.onSelectionChange();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Seçim değişimi.
  useEffect(() => {
    const onSel = () => {
      const el = ref.current;
      if (!el || composing.current) return;
      const s = domSelToSel(el);
      if (!s) return;
      if (cmpPos(s.anchor, selRef.current.anchor) !== 0 || cmpPos(s.focus, selRef.current.focus) !== 0) {
        selRef.current = s;
        pendingRef.current = null;
        propsRef.current.onSelectionChange();
      }
    };
    document.addEventListener("selectionchange", onSel);
    return () => document.removeEventListener("selectionchange", onSel);
  }, []);

  // Girdi olayları.
  useEffect(() => {
    const el = ref.current!;
    const insProps = () => pendingRef.current ?? propsAt(bodyRef.current, orderSel(selRef.current)[0]);
    const onBeforeInput = (e: InputEvent) => {
      if (composing.current || e.inputType === "insertCompositionText") return;
      e.preventDefault();
      const b = bodyRef.current;
      const sel = domSelToSel(el) ?? selRef.current;
      selRef.current = sel;
      const [from, to] = orderSel(sel);
      switch (e.inputType) {
        case "insertText":
        case "insertReplacementText": {
          let text = e.data ?? e.dataTransfer?.getData("text/plain") ?? "";
          let s = sel;
          if (e.inputType === "insertReplacementText") {
            const tr = e.getTargetRanges?.()[0];
            if (tr) {
              const a = domToPos(el, tr.startContainer, tr.startOffset);
              const f = domToPos(el, tr.endContainer, tr.endOffset);
              if (a && f) s = { anchor: a, focus: f };
            }
          }
          text = text.replace(/\n/g, "\v");
          const r = insertText(b, s, text, insProps());
          commit(r.body, caret(r.caret.p, r.caret.o), "text");
          break;
        }
        case "insertParagraph": {
          // Enter ayrı bir geri alma adımıdır.
          const r = splitParagraph(b, sel, insProps());
          commit(r.body, caret(r.caret.p, r.caret.o));
          break;
        }
        case "insertLineBreak": {
          const r = insertText(b, sel, "\v", insProps());
          commit(r.body, caret(r.caret.p, r.caret.o), "text");
          break;
        }
        case "deleteContentBackward":
        case "deleteContentForward":
        case "deleteWordBackward":
        case "deleteWordForward":
        case "deleteSoftLineBackward":
        case "deleteHardLineBackward":
        case "deleteSoftLineForward":
        case "deleteHardLineForward":
        case "deleteContent":
        case "deleteByCut": {
          if (e.inputType === "deleteContentBackward" || e.inputType === "deleteContentForward") {
            const r = e.inputType === "deleteContentBackward" ? removeBackward(b, sel) : removeForward(b, sel);
            if (r.body !== b) commit(r.body, caret(r.caret.p, r.caret.o), "del");
            break;
          }
          let a = from;
          let z = to;
          if (isCollapsed(sel)) {
            if (e.inputType === "deleteWordBackward") a = wordBoundary(b, from, -1);
            else if (e.inputType === "deleteWordForward") z = wordBoundary(b, from, 1);
            else if (e.inputType.endsWith("LineBackward")) a = { p: from.p, o: 0 };
            else if (e.inputType.endsWith("LineForward")) z = { p: from.p, o: paraLen(b.paragraphs[from.p]) };
          }
          if (cmpPos(a, z) === 0) return;
          const nb = deleteRange(b, a, z);
          commit(nb, caret(a.p, a.o), "del");
          break;
        }
        case "historyUndo":
          propsRef.current.onUndo();
          break;
        case "historyRedo":
          propsRef.current.onRedo();
          break;
        default:
          break;
      }
    };
    const onCompStart = () => {
      composing.current = true;
    };
    const onCompEnd = (e: CompositionEvent) => {
      composing.current = false;
      const data = e.data ?? "";
      render(true);
      if (data) {
        const r = insertText(bodyRef.current, selRef.current, data, insProps());
        commit(r.body, caret(r.caret.p, r.caret.o), "text");
      } else restoreSel();
    };
    const onCopy = (e: ClipboardEvent, cut: boolean) => {
      // selectionchange gecikebilir: kopyalanacak aralık her zaman DOM seçiminden alınır.
      const sel = (!composing.current && domSelToSel(el)) || selRef.current;
      selRef.current = sel;
      if (isCollapsed(sel)) return;
      e.preventDefault();
      const [a, z] = orderSel(sel);
      const paras = sliceBody(bodyRef.current, a, z);
      e.clipboardData?.setData("text/plain", textInRange(bodyRef.current, a, z).replace(/\v/g, "\n"));
      e.clipboardData?.setData(RICH_MIME, JSON.stringify(paras));
      propsRef.current.onCopyRich(paras);
      if (cut) commit(deleteRange(bodyRef.current, a, z), caret(a.p, a.o), "cut");
    };
    const onPaste = (e: ClipboardEvent) => {
      e.preventDefault();
      e.stopPropagation();
      const ds = domSelToSel(el);
      if (ds) selRef.current = ds;
      const rich = e.clipboardData?.getData(RICH_MIME);
      let paras: Paragraph[] | null = null;
      if (rich) {
        try {
          paras = JSON.parse(rich) as Paragraph[];
        } catch {
          paras = null;
        }
      }
      if (paras?.length) {
        const r = insertParagraphs(bodyRef.current, selRef.current, paras);
        commit(r.body, caret(r.caret.p, r.caret.o), "paste");
        return;
      }
      const text = e.clipboardData?.getData("text/plain") ?? "";
      if (!text) return;
      const r = insertText(bodyRef.current, selRef.current, text.replace(/\t/g, "\t"), insProps());
      commit(r.body, caret(r.caret.p, r.caret.o), "paste");
    };
    const copy = (e: ClipboardEvent) => onCopy(e, false);
    const cut = (e: ClipboardEvent) => onCopy(e, true);
    const onDrop = (e: DragEvent) => e.preventDefault();
    el.addEventListener("beforeinput", onBeforeInput);
    el.addEventListener("compositionstart", onCompStart);
    el.addEventListener("compositionend", onCompEnd);
    el.addEventListener("copy", copy);
    el.addEventListener("cut", cut);
    el.addEventListener("paste", onPaste);
    el.addEventListener("drop", onDrop);
    return () => {
      el.removeEventListener("beforeinput", onBeforeInput);
      el.removeEventListener("compositionstart", onCompStart);
      el.removeEventListener("compositionend", onCompEnd);
      el.removeEventListener("copy", copy);
      el.removeEventListener("cut", cut);
      el.removeEventListener("paste", onPaste);
      el.removeEventListener("drop", onDrop);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Düzenleyici tanıtıcısı (şerit komutları için).
  useEffect(() => {
    /** Tarayıcının selectionchange olayı gecikebilir: komuttan önce seçimi DOM'dan eşitle. */
    const sync = () => {
      const el = ref.current;
      if (!el || composing.current || document.activeElement !== el) return;
      const s = domSelToSel(el);
      if (s && (cmpPos(s.anchor, selRef.current.anchor) !== 0 || cmpPos(s.focus, selRef.current.focus) !== 0)) {
        selRef.current = s;
        pendingRef.current = null;
      }
    };
    const changeLevel = (delta: number) => {
      sync();
      const [a, z] = orderSel(selRef.current);
      const max = propsRef.current.maxLevel ?? 8;
      const nb = applyParaProps(bodyRef.current, a.p, z.p, (p) => ({ ...p, level: Math.max(0, Math.min(max, p.level + delta)) }));
      commit(nb, selRef.current);
    };
    const handle: TextEditorHandle = {
      applyRun(patch) {
        sync();
        const b = bodyRef.current;
        let [a, z] = orderSel(selRef.current);
        if (cmpPos(a, z) === 0) {
          const w = wordRangeAt(b, a);
          if (!w) {
            pendingRef.current = patch(pendingRef.current ?? propsAt(b, a), b.paragraphs[a.p]);
            propsRef.current.onSelectionChange();
            if (!b.paragraphs[a.p].runs.length) {
              commit(applyRunProps(b, a, z, patch), selRef.current, undefined);
              pendingRef.current = null;
            }
            return;
          }
          [a, z] = w;
        }
        const keep = selRef.current;
        commit(applyRunProps(b, a, z, patch), keep);
        propsRef.current.onSelectionChange();
      },
      applyPara(patch, levelDelta) {
        sync();
        if (levelDelta) return changeLevel(levelDelta);
        const [a, z] = orderSel(selRef.current);
        commit(applyParaProps(bodyRef.current, a.p, z.p, patch), selRef.current);
      },
      runProps() {
        sync();
        const b = bodyRef.current;
        const [a, z] = orderSel(selRef.current);
        const lvl = propsRef.current.ctx.lvl;
        if (pendingRef.current && cmpPos(a, z) === 0) return [effectiveRun(pendingRef.current, b.paragraphs[a.p], lvl)];
        const list = runPropsInRange(b, a, z);
        return list.map((rp) => effectiveRun(rp, b.paragraphs[Math.min(a.p, b.paragraphs.length - 1)], lvl));
      },
      paragraphs() {
        sync();
        const [a, z] = orderSel(selRef.current);
        return bodyRef.current.paragraphs.slice(a.p, z.p + 1);
      },
      transform(fn) {
        sync();
        let [a, z] = orderSel(selRef.current);
        if (cmpPos(a, z) === 0) {
          const w = wordRangeAt(bodyRef.current, a);
          if (!w) return;
          [a, z] = w;
        }
        commit(transformText(bodyRef.current, a, z, fn), selRef.current);
      },
      insert(text) {
        sync();
        const r = insertText(bodyRef.current, selRef.current, text, insProps());
        commit(r.body, caret(r.caret.p, r.caret.o), "text");
      },
      insertRich(paras) {
        sync();
        const r = insertParagraphs(bodyRef.current, selRef.current, paras);
        commit(r.body, caret(r.caret.p, r.caret.o), "paste");
      },
      insertRun(text, props) {
        sync();
        const r = insertText(bodyRef.current, selRef.current, text, { ...insProps(), ...props });
        commit(r.body, caret(r.caret.p, r.caret.o));
      },
      selectAll() {
        selRef.current = selectAll(bodyRef.current);
        restoreSel();
        propsRef.current.onSelectionChange();
      },
      selection: () => (sync(), selRef.current),
      setSelection(sel) {
        selRef.current = clampSel(sel);
        ref.current?.focus({ preventScroll: true });
        restoreSel();
        propsRef.current.onSelectionChange();
      },
      selectedText() {
        sync();
        const [a, z] = orderSel(selRef.current);
        return textInRange(bodyRef.current, a, z);
      },
      focus() {
        ref.current?.focus({ preventScroll: true });
        restoreSel();
      },
      placeCaret(x, y) {
        const el = ref.current;
        if (!el) return;
        const r = el.getBoundingClientRect();
        // Metin bloğunun dışındaki nokta bloğun kenarına çekilir.
        const cx = Math.max(r.left + 1, Math.min(r.right - 1, x));
        const cy = Math.max(r.top + 1, Math.min(r.bottom - 1, y));
        const b = bodyRef.current;
        const last = b.paragraphs.length - 1;
        const p = posFromPoint(el, cx, cy) ?? (y < r.top ? { p: 0, o: 0 } : { p: last, o: paraLen(b.paragraphs[last]) });
        selRef.current = { anchor: p, focus: p };
        el.focus({ preventScroll: true });
        restoreSel();
        propsRef.current.onSelectionChange();
      },
    };
    const insProps = () => pendingRef.current ?? propsAt(bodyRef.current, orderSel(selRef.current)[0]);
    propsRef.current.register(handle);
    return () => propsRef.current.register(null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      props.onExit();
      return;
    }
    if (e.key === "Tab" && !e.ctrlKey && !e.altKey) {
      e.preventDefault();
      e.stopPropagation();
      if (props.onTab?.(e.shiftKey)) return;
      const sel = selRef.current;
      const [a, z] = orderSel(sel);
      const b = bodyRef.current;
      const p = b.paragraphs[a.p];
      const atStart = a.o === 0 && cmpPos(a, z) === 0;
      const multi = a.p !== z.p;
      if (e.shiftKey || atStart || multi || (a.o === 0 && z.o === paraLen(p))) {
        const max = props.maxLevel ?? 8;
        const nb = applyParaProps(b, a.p, z.p, (pp) => ({ ...pp, level: Math.max(0, Math.min(max, pp.level + (e.shiftKey ? -1 : 1))) }));
        commit(nb, sel);
      } else {
        const r = insertText(b, sel, "\t", pendingRef.current ?? propsAt(b, a));
        commit(r.body, caret(r.caret.p, r.caret.o), "text");
      }
      return;
    }
    // Uygulama kısayolları (Ctrl+…) pencereye kabarcıklanır; ok tuşları vb. yerel kalır.
    if (!e.ctrlKey && !e.metaKey && (e.key.startsWith("Arrow") || e.key === "Delete" || e.key === "Backspace" || e.key === "Home" || e.key === "End" || e.key === "PageUp" || e.key === "PageDown" || e.key === "Enter")) {
      e.stopPropagation();
    }
  };

  return (
    <div
      ref={ref}
      className="tb-c tb-edit"
      contentEditable
      suppressContentEditableWarning
      spellCheck={props.spellcheck}
      lang={props.lang ?? "tr-TR"}
      onKeyDown={onKeyDown}
      onPointerDown={(e) => e.stopPropagation()}
      onDoubleClick={(e) => e.stopPropagation()}
    />
  );
}
