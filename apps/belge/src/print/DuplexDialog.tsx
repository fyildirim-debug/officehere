// Elle iki taraflı yazdırma: ön yüzlerden sonra kağıdın yeniden yerleştirilmesi istenir.
import { Dialog, DialogButton } from "@officehere/ui";

export function DuplexDialog({ open, onNext, onCancel }: { open: boolean; onNext: () => void; onCancel: () => void }) {
  if (!open) return null;
  return (
    <Dialog
      open
      title="Elle İki Taraflı Yazdır"
      width={420}
      onClose={onCancel}
      onSubmit={onNext}
      footer={
        <>
          <DialogButton primary onClick={onNext}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onCancel}>İptal</DialogButton>
        </>
      }
    >
      <p>Ön yüzler yazdırıldı. Yazdırılan kağıtları ters çevirip yazıcıya yeniden yerleştirin, ardından arka yüzleri yazdırmak için Tamam'a basın.</p>
    </Dialog>
  );
}
