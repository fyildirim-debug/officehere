/** Izgaranın üstündeki katmanlar: özet tablo yer tutucuları, yüzen nesneler, notlar, doğrulama ve toplam satırı listeleri, hata etiketi, çizim kipi. */
import type { Controller } from "../controller";
import { FloatingLayer } from "../charts/FloatingLayer";
import { DrawOverlay } from "../charts/DrawOverlay";
import { SlicerLayer } from "../charts/SlicerLayer";
import { PivotPlaceholders } from "../pivot/PivotPlaceholder";
import { AllNotes, ValidationUi } from "./ValidationUi";
import { TotalRowUi } from "./TotalRowUi";
import { ErrorTag } from "./ErrorTag";
import type { GridLayout } from "./layout";
import { FlashPreviewUi } from "./FlashPreviewUi";

export function GridOverlays({ ctl, layout }: { ctl: Controller; layout: GridLayout }) {
  return (
    <>
      <PivotPlaceholders ctl={ctl} layout={layout} />
      <FloatingLayer ctl={ctl} layout={layout} />
      <SlicerLayer ctl={ctl} layout={layout} />
      {ctl.showAllNotes && <AllNotes ctl={ctl} layout={layout} />}
      <ValidationUi ctl={ctl} layout={layout} />
      <TotalRowUi ctl={ctl} layout={layout} />
      <ErrorTag ctl={ctl} layout={layout} />
      <DrawOverlay ctl={ctl} layout={layout} />
      {ctl.edit && <FlashPreviewUi ctl={ctl} layout={layout} />}
    </>
  );
}
