// Başvurular sekmesi grupları: Alıntılar ve Kaynakça, Resim Yazıları.
import { ArrowSync20Regular, BookAdd24Regular, BookDatabase32Regular, BookOpen32Regular, ImageTable24Regular, LinkSquare24Regular, TextQuote24Regular } from "@fluentui/react-icons";
import { Menu, MenuHeader, MenuItem, MenuSeparator, RibbonButton, RibbonGroup, RibbonMenuButton, RibbonStack } from "@officehere/ui";
import { useBelge } from "../state/context";
import { BIB_STYLE_LABEL, citationText, type BibStyle } from "../model/sources";
import { setFieldContext } from "../editor/fields";

export function CitationsGroup() {
  const ctx = useBelge();
  const sources = ctx.settings.sources ?? [];
  const style = ctx.settings.bibStyle ?? "apa";
  const ed = ctx.mainEditor;
  const cite = (tag: string) => ed.chain().focus().insertContent({ type: "citation", attrs: { tag } }).run();
  const insertBib = (title: string) => {
    ed.chain()
      .focus()
      .insertContent({ type: "bibliography", attrs: { title, entries: [] } })
      .run();
    setFieldContext(ed.view, {});
  };
  return (
    <RibbonGroup label="Alıntılar ve Kaynakça" icon={<BookOpen32Regular />}>
      <RibbonMenuButton
        size="large"
        icon={<TextQuote24Regular />}
        label="Alıntı Ekle"
        disabled={!ctx.inBody}
        menu={
          <Menu width={300}>
            {sources.length > 0 && <MenuHeader>Kaynaklar</MenuHeader>}
            {sources.map((s, i) => (
              <MenuItem key={s.tag} label={`${s.author || s.title}`} description={`${s.title} ${citationText(s, style, i)}`} onClick={() => cite(s.tag)} />
            ))}
            {sources.length > 0 && <MenuSeparator />}
            <MenuItem icon={<BookAdd24Regular />} label="Yeni Kaynak Ekle..." onClick={() => ctx.openDialog("source", "insert")} />
          </Menu>
        }
      />
      <RibbonStack>
        <RibbonButton icon={<BookDatabase32Regular />} label="Kaynakları Yönet" onClick={() => ctx.openDialog("sources")} />
        <RibbonMenuButton
          label={`Stil: ${BIB_STYLE_LABEL[style]}`}
          tip="Kaynakça Stili"
          menu={
            <Menu>
              {(Object.keys(BIB_STYLE_LABEL) as BibStyle[]).map((k) => (
                <MenuItem
                  key={k}
                  label={BIB_STYLE_LABEL[k]}
                  description={k === "apa" ? "Yazar-tarih (Soyadı, Yıl)" : "Numaralı [1]"}
                  checked={style === k}
                  onClick={() => ctx.updateSettings((s) => ({ ...s, bibStyle: k }))}
                />
              ))}
            </Menu>
          }
        />
        <RibbonMenuButton
          icon={<BookOpen32Regular />}
          label="Kaynakça"
          disabled={!ctx.inBody}
          menu={
            <Menu width={240}>
              <MenuHeader>Yerleşik</MenuHeader>
              <MenuItem label="Kaynakça" onClick={() => insertBib("Kaynakça")} />
              <MenuItem label="Başvurular" onClick={() => insertBib("Başvurular")} />
              <MenuItem label="Başvurulan Eserler" onClick={() => insertBib("Başvurulan Eserler")} />
              <MenuSeparator />
              <MenuItem icon={<ArrowSync20Regular />} label="Kaynakçayı Güncelleştir" onClick={() => setFieldContext(ed.view, {})} />
            </Menu>
          }
        />
      </RibbonStack>
    </RibbonGroup>
  );
}

export function CaptionsGroup() {
  const ctx = useBelge();
  return (
    <RibbonGroup label="Resim Yazıları" icon={<ImageTable24Regular />}>
      <RibbonButton size="large" icon={<ImageTable24Regular />} label="Resim Yazısı Ekle" disabled={!ctx.inBody} onClick={() => ctx.openDialog("caption")} />
      <RibbonStack>
        <RibbonButton icon={<ImageTable24Regular />} label="Şekiller Tablosu Ekle" disabled={!ctx.inBody} onClick={() => ctx.openDialog("tableOfFigures")} />
        <RibbonButton icon={<ArrowSync20Regular />} label="Tabloyu Güncelleştir" onClick={ctx.updateToc} />
        <RibbonButton icon={<LinkSquare24Regular />} label="Çapraz Başvuru" disabled={!ctx.inBody} onClick={() => ctx.openDialog("crossRef")} />
      </RibbonStack>
    </RibbonGroup>
  );
}
