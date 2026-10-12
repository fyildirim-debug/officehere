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
    return () => dom.removeEventListener("click", onClick);
  }, [editor, openLink]);

  return openLink;
}
