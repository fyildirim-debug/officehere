// Bul ve Değiştir iletişim kutusu (Bul / Değiştir / Git sekmeleri).
import { useEffect, useState } from "react";
import { TextSelection } from "@tiptap/pm/state";
import { Dialog, DialogButton, DialogTabs, Field } from "@officehere/ui";
import { useBelge } from "../state/context";
import { getSearch, gotoMatch, replaceAll, replaceCurrent, setSearch } from "../editor/search";
import { Check } from "./FontDialog";
import { AutoFocus } from "./AutoFocus";

export function FindReplaceDialog({ initialTab, onClose }: { initialTab: string; onClose: () => void }) {
  const { editor, controller, notify } = useBelge();
  const [tab, setTab] = useState(initialTab === "goto" ? "goto" : initialTab === "replace" ? "replace" : "find");
  const st = getSearch(editor.state);
  const selText = editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to, " ");
  const [query, setQuery] = useState(selText && selText.length < 80 ? selText : (st?.query ?? ""));
  const [repl, setRepl] = useState("");
  const [matchCase, setMatchCase] = useState(st?.caseSensitive ?? false);
  const [whole, setWhole] = useState(st?.wholeWord ?? false);
  const [more, setMore] = useState(false);
  const [gotoKind, setGotoKind] = useState("page");
  const [gotoVal, setGotoVal] = useState("");
  const [msg, setMsg] = useState("");

  useEffect(() => {
    setSearch(editor.view, { query, caseSensitive: matchCase, wholeWord: whole });
    setMsg("");
  }, [editor, query, matchCase, whole]);

  const findNext = () => {
    if (!gotoMatch(editor.view, 1)) setMsg("Arama öğesi bulunamadı.");
    else setMsg("");
  };

  const doReplaceAll = () => {
    const n = replaceAll(editor.view, repl);
    notify(n ? `Tümü tamam. ${n} değiştirme yapıldı.` : "Arama öğesi bulunamadı.");
    setMsg(n ? `${n} değiştirme yapıldı.` : "Arama öğesi bulunamadı.");
  };

  const doGoto = () => {
    const n = parseInt(gotoVal, 10);
    if (gotoKind === "page") {
      const info = controller.info;
      const idx = Number.isFinite(n) ? Math.min(info.pageCount, Math.max(1, n)) - 1 : Math.min(info.pageCount - 1, controller.pageOf(editor.state.selection.from) + 1);
      const pos = info.pageStarts[idx] ?? 0;
      editor.view.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(Math.min(pos + 1, editor.state.doc.content.size)))).scrollIntoView());
    } else if (gotoKind === "line" || gotoKind === "heading") {
      const targets: number[] = [];
      editor.state.doc.descendants((node, pos) => {
        if (gotoKind === "heading" ? node.type.name === "heading" : node.isTextblock) {
          targets.push(pos);
          return false;
        }
        return true;
      });
      const idx = Number.isFinite(n) ? Math.min(targets.length, Math.max(1, n)) - 1 : 0;
      if (targets[idx] !== undefined) editor.view.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(targets[idx] + 1))).scrollIntoView());
    }
    editor.commands.focus();
  };

  const footer =
    tab === "goto" ? (
      <>
        <DialogButton primary onClick={doGoto}>
          {gotoVal ? "Git" : "Sonraki"}
        </DialogButton>
        <DialogButton onClick={onClose}>Kapat</DialogButton>
      </>
    ) : (
      <>
        <DialogButton onClick={() => setMore(!more)}>{more ? "<< Daha Az" : "Diğer >>"}</DialogButton>
        <span className="bl-flex" />
        {tab === "replace" && (
          <>
            <DialogButton onClick={() => replaceCurrent(editor.view, repl)}>Değiştir</DialogButton>
            <DialogButton onClick={doReplaceAll}>Tümünü Değiştir</DialogButton>
          </>
        )}
        <DialogButton primary onClick={findNext}>
          Sonrakini Bul
        </DialogButton>
        <DialogButton onClick={onClose}>İptal</DialogButton>
      </>
    );

  return (
    <Dialog open title="Bul ve Değiştir" width={580} onClose={onClose} onSubmit={tab === "goto" ? doGoto : findNext} footer={footer}>
      <DialogTabs
        tabs={[
          { id: "find", label: "Bul" },
          { id: "replace", label: "Değiştir" },
          { id: "goto", label: "Git" },
        ]}
        active={tab}
        onChange={setTab}
      />
      <AutoFocus deps={[tab]}>
      {tab !== "goto" ? (
        <div className="bl-find-fields">
          <Field label="Aranan:" inline>
            <input type="text" value={query} onChange={(e) => setQuery(e.target.value)} style={{ flex: 1 }} />
          </Field>
          {tab === "replace" && (
            <Field label="Yeni değer:" inline>
              <input type="text" value={repl} onChange={(e) => setRepl(e.target.value)} style={{ flex: 1 }} />
            </Field>
          )}
          <div className="bl-dlg-msg">{msg || (query ? `${getSearch(editor.state)?.matches.length ?? 0} eşleşme` : "")}</div>
          {more && (
            <div className="bl-dlg-checks is-single">
              <Check label="Büyük/küçük harf eşleştir" v={matchCase} set={setMatchCase} />
              <Check label="Yalnızca tam sözcükleri bul" v={whole} set={setWhole} />
            </div>
          )}
        </div>
      ) : (
        <div className="bl-dlg-goto">
          <div className="bl-listfield">
            <span>Git:</span>
            <select size={4} value={gotoKind} onChange={(e) => setGotoKind(e.target.value)}>
              <option value="page">Sayfa</option>
              <option value="line">Paragraf</option>
              <option value="heading">Başlık</option>
            </select>
          </div>
          <Field label={gotoKind === "page" ? "Sayfa numarasını girin:" : gotoKind === "heading" ? "Başlık numarasını girin:" : "Paragraf numarasını girin:"}>
            <input type="text" value={gotoVal} onChange={(e) => setGotoVal(e.target.value)} />
          </Field>
        </div>
      )}
      </AutoFocus>
    </Dialog>
  );
}
