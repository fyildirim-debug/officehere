// Belgedeki tıklamalar: Ctrl+tıklama ile köprüyü ya da İçindekiler girdisini izleme; yer işaretine gitme.
import { useCallback, useEffect } from "react";
import type { Editor } from "@tiptap/core";
import { TextSelection } from "@tiptap/pm/state";
import { collectHeadings } from "../editor/commands";

export function useEditorClicks(editor: Editor) {
  const openLink = useCallback(
    (href: string) => {
      if (href.startsWith("#")) {
        const name = href.slice(1);
        let target = -1;
        editor.state.doc.descendants((n, pos) => {
          if (target < 0 && n.type.name === "bookmark" && n.attrs.name === name) target = pos;
          return target < 0;
        });
        if (name === "_top") target = 0;
        if (target >= 0) editor.view.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(target))).scrollIntoView());
        return;
      }
      window.open(href, "_blank", "noopener");
    },
    [editor],
  );

  useEffect(() => {
    const dom = editor.view.dom;
    const onClick = (e: MouseEvent) => {
      const a = (e.target as HTMLElement).closest("a[href]");
      if (a && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        openLink(a.getAttribute("href") ?? "");
      } else if (a) e.preventDefault();
      const toc = (e.target as HTMLElement).closest(".bl-toc-entry");
      if (toc && (e.ctrlKey || e.metaKey)) {
        const pos = Number(toc.getAttribute("data-pos"));
        const all = collectHeadings(editor.state.doc);
        const heading = all.find((h) => h.pos === pos) ?? all.find((h) => h.text === toc.querySelector(".bl-toc-text")?.textContent);
        if (heading) editor.view.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(heading.pos + 1))).scrollIntoView());
      }
    };
    dom.addEventListener("click", onClick);
    // Köprü üzerinde ipucu: adres ve "Ctrl tuşuna basıp tıklayın" (Word'deki gibi).
    let tip: HTMLDivElement | null = null;
    let timer = 0;
    const hide = () => {
      window.clearTimeout(timer);
      tip?.remove();
      tip = null;
    };
    const onOver = (e: MouseEvent) => {
      const a = (e.target as HTMLElement).closest("a[href], .bl-toc-entry");
      if (!a) return hide();
      if (tip?.dataset.for === (a.getAttribute("href") ?? "toc")) return;
      hide();
      timer = window.setTimeout(() => {
        const href = a.getAttribute("href");
        tip = document.createElement("div");
        tip.className = "bl-linktip";
        tip.setAttribute("role", "tooltip");
        tip.dataset.for = href ?? "toc";
        if (href) {
          const u = document.createElement("div");
          u.className = "bl-linktip-url";
          u.textContent = href.startsWith("#") ? `Yer işareti: ${href.slice(1)}` : href;
          tip.appendChild(u);
        }
        const hint = document.createElement("div");
        hint.textContent = "Bağlantıyı izlemek için Ctrl tuşuna basıp tıklayın";
        tip.appendChild(hint);
        document.body.appendChild(tip);
        const r = a.getBoundingClientRect();
        const w = tip.offsetWidth;
        tip.style.left = `${Math.max(4, Math.min(window.innerWidth - w - 4, e.clientX - w / 2))}px`;
        tip.style.top = `${r.top - tip.offsetHeight - 6 < 4 ? r.bottom + 6 : r.top - tip.offsetHeight - 6}px`;
      }, 350);
    };
    dom.addEventListener("mouseover", onOver);
    dom.addEventListener("mouseleave", hide);
    dom.addEventListener("keydown", hide);
    return () => {
      dom.removeEventListener("click", onClick);
      dom.removeEventListener("mouseover", onOver);
      dom.removeEventListener("mouseleave", hide);
      dom.removeEventListener("keydown", hide);
      hide();
    };
  }, [editor, openLink]);

  return openLink;
}
