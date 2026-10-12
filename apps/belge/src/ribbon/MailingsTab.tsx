// Postalar sekmesi: Adres Mektup Birleştirmeyi Başlat (Alıcıları Seç, Alıcı Listesini Düzenle), Alanları Yaz
// ve Ekle (Birleştirme Alanı Ekle, Birleştirme Alanlarını Vurgula), Sonuçları Önizle, Son ve Birleştir.
import { useRef } from "react";
import {
  ArrowNext24Regular,
  ArrowPrevious24Regular,
  ArrowLeft24Regular,
  ArrowRight24Regular,
  Eye24Regular,
  Highlight20Regular,
  MailMultiple32Regular,
  Merge24Regular,
  PeopleList24Regular,
  Textbox24Regular,
} from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator, RibbonButton, RibbonGroup, RibbonMenuButton, RibbonStack } from "@officehere/ui";
import { useBelge } from "../state/context";
import { useMerge } from "../mailmerge/useMailMerge";
import { readRecipients } from "../mailmerge/data";

export function MailingsTab() {
  const ctx = useBelge();
  const m = useMerge();
  const input = useRef<HTMLInputElement>(null);
  const fields = m.list?.fields ?? [];
  const count = m.list?.records.length ?? 0;
  const insertField = (name: string) => ctx.mainEditor.chain().focus().insertContent({ type: "mergeField", attrs: { name } }).run();
  const onFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    e.target.value = "";
    if (!f) return;
    try {
      const list = await readRecipients(f.name, await f.arrayBuffer());
      if (!list.records.length) return ctx.notify(`"${f.name}" içinde alıcı kaydı bulunamadı.`, "warning");
      m.setList(list);
      ctx.notify(`${list.records.length} alıcı yüklendi (${list.fields.length} alan).`, "success");
    } catch (err) {
      ctx.notify(`Alıcı listesi okunamadı: ${(err as Error).message}`, "error");
    }
  };
  return (
    <div className="bl-tab-mailings">
      <RibbonGroup label="Adres Mektup Birleştirmeyi Başlat" icon={<MailMultiple32Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<MailMultiple32Regular />}
          label="Adres Mektup Birleştirmeyi Başlat"
          menu={
            <Menu width={260}>
              <MenuItem label="Mektuplar" checked onClick={() => {}} />
              <MenuItem label="Normal Belge" onClick={() => m.setList(null)} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<PeopleList24Regular />}
          label="Alıcıları Seç"
          menu={
            <Menu width={260}>
              <MenuItem label="Var Olan Listeyi Kullan..." description="CSV ya da XLSX dosyası" onClick={() => input.current?.click()} />
              {m.list && (
                <>
                  <MenuSeparator />
                  <MenuItem label="Listeyi Kaldır" onClick={() => m.setList(null)} />
                </>
              )}
            </Menu>
          }
        />
        <RibbonButton size="large" icon={<PeopleList24Regular />} label="Alıcı Listesini Düzenle" disabled={!m.list} onClick={() => ctx.openDialog("recipients")} />
      </RibbonGroup>
      <RibbonGroup label="Alanları Yaz ve Ekle" icon={<Textbox24Regular />}>
        <RibbonStack>
          <RibbonButton icon={<Highlight20Regular />} label="Birleştirme Alanlarını Vurgula" checked={m.highlight} onClick={() => m.setHighlight(!m.highlight)} />
        </RibbonStack>
        <RibbonMenuButton
          size="large"
          icon={<Textbox24Regular />}
          label="Birleştirme Alanı Ekle"
          disabled={!m.list || !ctx.inBody}
          menu={
            <Menu width={220}>
              {fields.map((f) => (
                <MenuItem key={f} label={f} onClick={() => insertField(f)} />
              ))}
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Sonuçları Önizle" icon={<Eye24Regular />}>
        <RibbonButton size="large" icon={<Eye24Regular />} label="Sonuçları Önizle" disabled={!m.list} checked={m.preview} onClick={() => m.setPreview(!m.preview)} />
        <RibbonStack>
          <div className="bl-merge-nav">
            <RibbonButton size="small" icon={<ArrowPrevious24Regular />} tip="İlk Kayıt" disabled={!m.preview || m.index <= 0} onClick={() => m.setIndex(0)} />
            <RibbonButton size="small" icon={<ArrowLeft24Regular />} tip="Önceki Kayıt" disabled={!m.preview || m.index <= 0} onClick={() => m.setIndex(m.index - 1)} />
            <input
              className="bl-merge-index"
              aria-label="Kayıt"
              type="number"
              min={1}
              max={count}
              disabled={!m.preview}
              value={m.preview ? m.index + 1 : ""}
              onChange={(e) => m.setIndex(Number(e.target.value) - 1)}
            />
            <RibbonButton size="small" icon={<ArrowRight24Regular />} tip="Sonraki Kayıt" disabled={!m.preview || m.index >= count - 1} onClick={() => m.setIndex(m.index + 1)} />
            <RibbonButton size="small" icon={<ArrowNext24Regular />} tip="Son Kayıt" disabled={!m.preview || m.index >= count - 1} onClick={() => m.setIndex(count - 1)} />
          </div>
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Son" icon={<Merge24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Merge24Regular />}
          label="Son ve Birleştir"
          disabled={!m.list}
          menu={
            <Menu width={240}>
              <MenuItem label="Tek Tek Belgeleri Düzenle..." onClick={() => ctx.openDialog("merge", "edit")} />
              <MenuItem label="Belgeleri Yazdır..." onClick={() => ctx.openDialog("merge", "print")} />
            </Menu>
          }
        />
      </RibbonGroup>
      <input ref={input} type="file" accept=".csv,.txt,.xlsx" className="oh-hidden-input" aria-label="Alıcı listesi" data-testid="bl-merge-input" onChange={onFile} />
    </div>
  );
}
