// Denklem iletişim kutusu: doğrusal (LaTeX benzeri) giriş, canlı önizleme, yapı/simge paletleri,
// satır içi / görüntülü seçimi. Seçili denklem varsa düzenlenir.
import { useMemo, useRef, useState } from "react";
import { NodeSelection } from "@tiptap/pm/state";
import { Dialog, DialogButton } from "@officehere/ui";
import { useBelge } from "../state/context";
import { renderEquation } from "../editor/equation";

/** Yapılar: eklenecek parça ve imlecin konacağı yer (¦ işareti). */
const STRUCTURES: { label: string; snippet: string }[] = [
  { label: "Kesir", snippet: "\\frac{¦}{}" },
  { label: "Üst Simge", snippet: "^{¦}" },
  { label: "Alt Simge", snippet: "_{¦}" },
  { label: "Kök", snippet: "\\sqrt{¦}" },
  { label: "n. Kök", snippet: "\\sqrt[¦]{}" },
  { label: "İntegral", snippet: "\\int_{¦}^{} " },
  { label: "Toplam", snippet: "\\sum_{¦}^{} " },
  { label: "Çarpım", snippet: "\\prod_{¦}^{} " },
  { label: "Ayraç", snippet: "\\left(¦\\right)" },
  { label: "Limit", snippet: "\\lim_{¦} " },
  { label: "İşlev (sin)", snippet: "\\sin ¦" },
  { label: "Vurgu", snippet: "\\hat{¦}" },
  { label: "Matris 2×2", snippet: "\\begin{pmatrix}¦&\\\\&\\end{pmatrix}" },
  { label: "Metin", snippet: "\\text{¦}" },
];

const SYMBOLS: [string, string][] = [
  ["±", "\\pm "], ["×", "\\times "], ["÷", "\\div "], ["⋅", "\\cdot "], ["≤", "\\leq "], ["≥", "\\geq "], ["≠", "\\neq "], ["≈", "\\approx "], ["≡", "\\equiv "], ["∞", "\\infty "],
  ["→", "\\to "], ["⇒", "\\Rightarrow "], ["∈", "\\in "], ["⊂", "\\subset "], ["∪", "\\cup "], ["∩", "\\cap "], ["∀", "\\forall "], ["∃", "\\exists "], ["∂", "\\partial "], ["∇", "\\nabla "],
  ["α", "\\alpha "], ["β", "\\beta "], ["γ", "\\gamma "], ["δ", "\\delta "], ["ε", "\\epsilon "], ["θ", "\\theta "], ["λ", "\\lambda "], ["μ", "\\mu "], ["π", "\\pi "], ["σ", "\\sigma "],
  ["φ", "\\phi "], ["ω", "\\omega "], ["Δ", "\\Delta "], ["Σ", "\\Sigma "], ["Ω", "\\Omega "], ["°", "\\degree "], ["…", "\\ldots "], ["′", "'"],
];

export function EquationDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const sel = editor.state.selection;
  const editing = sel instanceof NodeSelection && sel.node.type.name === "equation" ? sel.node : null;
  const [latex, setLatex] = useState<string>(editing ? String(editing.attrs.latex) : "");
  const emptyPara = sel.$from.parent.isTextblock && sel.$from.parent.content.size === 0;
  const [display, setDisplay] = useState<boolean>(editing ? !!editing.attrs.display : emptyPara);
  const area = useRef<HTMLTextAreaElement>(null);
  // Önizleme okunaklı olsun diye her zaman görüntülü biçimde çizilir.
  const preview = useMemo(() => (latex.trim() ? renderEquation(latex, true) : ""), [latex]);

  const put = (snippet: string) => {
    const ta = area.current;
    const at = ta ? ta.selectionStart : latex.length;
    const end = ta ? ta.selectionEnd : latex.length;
    const caret = snippet.indexOf("¦");
    const text = snippet.replace("¦", "");
    const next = latex.slice(0, at) + text + latex.slice(end);
    setLatex(next);
    requestAnimationFrame(() => {
      if (!ta) return;
      ta.focus();
      const c = at + (caret >= 0 ? caret : text.length);
      ta.setSelectionRange(c, c);
    });
  };

  const apply = () => {
    if (!latex.trim()) return onClose();
    editor.chain().focus().insertEquation(latex.trim(), display).run();
    onClose();
  };

  return (
    <Dialog
      open
      title="Denklem"
      width={620}
      onClose={onClose}
      footer={
        <>
          <DialogButton primary onClick={apply} disabled={!latex.trim()}>
            {editing ? "Güncelleştir" : "Ekle"}
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <div className="bl-eq-palette" role="toolbar" aria-label="Yapılar">
        {STRUCTURES.map((s) => (
          <button key={s.label} type="button" className="oh-btn" onClick={() => put(s.snippet)}>
            {s.label}
          </button>
        ))}
      </div>
      <div className="bl-eq-symbols" role="toolbar" aria-label="Simgeler">
        {SYMBOLS.map(([ch, code]) => (
          <button key={code} type="button" title={code.trim()} onClick={() => put(code)}>
            {ch}
          </button>
        ))}
      </div>
      <label className="bl-eq-label" htmlFor="bl-eq-input">
        {"Doğrusal biçim (ör. x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a} ya da a/b, x^2, \\alpha):"}
      </label>
      <textarea
        id="bl-eq-input"
        ref={area}
        className="bl-eq-input"
        rows={3}
        value={latex}
        spellCheck={false}
        aria-label="Denklem"
        autoFocus
        onChange={(e) => setLatex(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            apply();
          }
        }}
      />
      <div className="bl-eq-preview bl-doc" aria-label="Önizleme" dangerouslySetInnerHTML={{ __html: preview || "<span class='bl-eq-empty'>Denkleminizi buraya yazın.</span>" }} />
      <div className="bl-eq-mode">
        <label className="bl-radio">
          <input type="radio" name="eqmode" checked={!display} onChange={() => setDisplay(false)} /> Satır İçi
        </label>
        <label className="bl-radio">
          <input type="radio" name="eqmode" checked={display} onChange={() => setDisplay(true)} /> Görüntü
        </label>
      </div>
    </Dialog>
  );
}
