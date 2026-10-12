/** Gözden Geçir sekmesi: Yazım Denetleme (istatistik), Notlar, Koruma. */
import { ShieldLock24Regular, TextProofingTools24Regular, ArrowLeft20Regular, ArrowRight20Regular, CommentNote20Regular, Delete20Regular, LockClosed24Regular, LockOpen24Regular, NoteAdd24Regular, NoteEdit24Regular, TextNumberFormat24Regular } from "@fluentui/react-icons";
import { RibbonButton, RibbonGroup, RibbonStack } from "@officehere/ui";
import type { Controller } from "../controller";
import { gotoNote } from "../actions/view";

export function ReviewTab({ ctl }: { ctl: Controller }) {
  const sheet = ctl.sheet;
  const { row, col } = ctl.sel.active;
  const note = sheet.noteAt(row, col);
  const anyNotes = ctl.wb.sheets.some((s) => s.notes.size > 0);
  return (
    <>
      <RibbonGroup label="Yazım Denetleme" icon={<TextNumberFormat24Regular />}>
        <RibbonButton
          size="large"
          icon={<TextProofingTools24Regular />}
          label="Yazım Denetimi"
          shortcut="F7"
          description="Sayfadaki metinleri sırayla denetler; hatalı sözcüklerin altı sistemin Türkçe yazım denetleyicisiyle çizilir."
          onClick={() => ctl.openDialog("spelling")}
        />
        <RibbonButton size="large" icon={<TextNumberFormat24Regular />} label="Çalışma Kitabı İstatistikleri" description="Geçerli sayfa ve çalışma kitabı hakkında istatistikleri gösterir." onClick={() => ctl.openDialog("workbookStats")} />
      </RibbonGroup>
      <RibbonGroup label="Notlar" icon={<NoteAdd24Regular />}>
        <RibbonButton size="large" icon={note ? <NoteEdit24Regular /> : <NoteAdd24Regular />} label={note ? "Notu Düzenle" : "Yeni Not"} shortcut="Shift+F2" onClick={() => ctl.openDialog("note")} />
        <RibbonStack>
          <RibbonButton icon={<Delete20Regular />} label="Sil" disabled={!note} onClick={() => ctl.ops.setNote(sheet, row, col, null)} />
          <RibbonButton icon={<ArrowLeft20Regular />} label="Önceki Not" disabled={!anyNotes} onClick={() => gotoNote(ctl, -1)} />
          <RibbonButton icon={<ArrowRight20Regular />} label="Sonraki Not" disabled={!anyNotes} onClick={() => gotoNote(ctl, 1)} />
        </RibbonStack>
        <RibbonStack>
          <RibbonButton
            icon={<CommentNote20Regular />}
            label="Tüm Notları Göster"
            checked={ctl.showAllNotes}
            onClick={() => {
              ctl.showAllNotes = !ctl.showAllNotes;
              ctl.emit();
            }}
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Koruma" icon={<LockClosed24Regular />}>
        <RibbonButton
          size="large"
          icon={sheet.protected ? <LockOpen24Regular /> : <LockClosed24Regular />}
          label={sheet.protected ? "Sayfa Korumasını Kaldır" : "Sayfayı Koru"}
          description="Kilitli hücrelerde istenmeyen değişiklikleri önler."
          onClick={() => ctl.openDialog("protect")}
        />
        <RibbonButton
          size="large"
          icon={<ShieldLock24Regular />}
          label={ctl.wb.protection?.structure ? "Çalışma Kitabı Korumasını Kaldır" : "Çalışma Kitabını Koru"}
          checked={!!ctl.wb.protection?.structure}
          description="Sayfaların eklenmesi, silinmesi, taşınması, gizlenmesi ve yeniden adlandırılmasını önler."
          onClick={() => ctl.openDialog("protectWorkbook")}
        />
      </RibbonGroup>
    </>
  );
}
