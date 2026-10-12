/** Veri sekmesi: Veri Al, Sırala ve Filtre Uygula, Veri Araçları. */
import {
  ArrowSync20Regular,
  ArrowSync24Regular,
  DocumentText24Regular,
  Filter24Regular,
  FilterDismiss20Regular,
  TextSortAscending20Regular,
  TextSortDescending20Regular,
  ArrowSort24Regular,
  TextColumnTwo24Regular,
  TableDismiss24Regular,
  CheckboxChecked24Regular,
} from "@fluentui/react-icons";
import { openFile } from "@officehere/core";
import { Menu, MenuItem, RibbonButton, RibbonGroup, RibbonSplitButton, RibbonStack } from "@officehere/ui";
import type { Controller } from "../controller";
import * as D from "../actions/data";
import { refreshAll } from "../pivot/pivotOps";

export async function importText(ctl: Controller): Promise<void> {
  const f = await openFile([{ name: "Metin dosyaları", extensions: ["csv", "txt", "tsv"] }]);
  if (f) ctl.openDialog("textImport", { name: f.name, data: f.data });
}

export function DataTab({ ctl }: { ctl: Controller }) {
  const sheet = ctl.sheet;
  const filters = D.filtersOf(sheet);
  const filterOn = !!sheet.autoFilter || !!sheet.tableAt(ctl.sel.active.row, ctl.sel.active.col)?.filterButton;
  const anyCriteria = filters.some((f) => Object.keys(f.columns).length > 0);
  return (
    <>
      <RibbonGroup label="Veri Al ve Dönüştür" icon={<DocumentText24Regular />}>
        <RibbonButton size="large" icon={<DocumentText24Regular />} label="Metinden/ CSV'den" description="Metin, virgülle ya da noktalı virgülle ayrılmış değer dosyasından veri alır." onClick={() => void importText(ctl)} />
      </RibbonGroup>
      <RibbonGroup label="Sorgular ve Bağlantılar" icon={<ArrowSync24Regular />}>
        <RibbonButton size="large" icon={<ArrowSync24Regular />} label="Tümünü Yenile" shortcut="Ctrl+Alt+F5" description="Çalışma kitabındaki tüm özet tabloları kaynak verilerinden yeniden hesaplar." disabled={!ctl.wb.sheets.some((s) => s.pivots.length)} onClick={() => refreshAll(ctl)} />
      </RibbonGroup>
      <RibbonGroup label="Sırala ve Filtre Uygula" icon={<Filter24Regular />}>
        <RibbonStack>
          <RibbonButton size="small" icon={<TextSortAscending20Regular />} tip="A'dan Z'ye Sırala" description="En düşükten en yükseğe sıralar." onClick={() => D.quickSort(ctl, false)} />
          <RibbonButton size="small" icon={<TextSortDescending20Regular />} tip="Z'den A'ya Sırala" description="En yüksekten en düşüğe sıralar." onClick={() => D.quickSort(ctl, true)} />
        </RibbonStack>
        <RibbonButton size="large" icon={<ArrowSort24Regular />} label="Sırala" description="Verileri birden çok ölçüte göre sıralar." onClick={() => ctl.openDialog("sort")} />
        <RibbonButton size="large" icon={<Filter24Regular />} label="Filtre" shortcut="Ctrl+Shift+L" checked={filterOn} description="Seçili hücreler için filtrelemeyi açar." onClick={() => D.toggleAutoFilter(ctl)} />
        <RibbonStack>
          <RibbonButton icon={<FilterDismiss20Regular />} label="Temizle" disabled={!anyCriteria} onClick={() => D.clearAllFilters(ctl)} />
          <RibbonButton icon={<ArrowSync20Regular />} label="Yeniden Uygula" disabled={!filters.length} onClick={() => D.reapplyFilters(ctl)} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Veri Araçları" icon={<CheckboxChecked24Regular />}>
        <RibbonButton size="large" icon={<TextColumnTwo24Regular />} label="Metni Sütunlara Dönüştür" className="tb-wide" description="Tek bir sütundaki metni ayırıcıya göre birden çok sütuna böler." onClick={() => ctl.openDialog("textToColumns")} />
        <RibbonButton size="large" icon={<TableDismiss24Regular />} label="Yinelenenleri Kaldır" description="Sayfadaki yinelenen satırları siler." onClick={() => ctl.openDialog("removeDuplicates")} />
        <RibbonSplitButton
          size="large"
          icon={<CheckboxChecked24Regular />}
          label="Veri Doğrulama"
          description="Hücreye girilebilecek veri türünü sınırlar (ör. açılır liste)."
          onClick={() => ctl.openDialog("dataValidation")}
          menu={
            <Menu width={220}>
              <MenuItem label="Veri Doğrulama..." onClick={() => ctl.openDialog("dataValidation")} />
              <MenuItem
                label="Doğrulamayı Temizle"
                onClick={() =>
                  ctl.ops.sheetOp(sheet, "Veri Doğrulama", () => {
                    const sel = ctl.sel.ranges;
                    sheet.validations = sheet.validations.filter((v) => !v.ranges.some((r) => sel.some((s) => s.r1 <= r.r2 && r.r1 <= s.r2 && s.c1 <= r.c2 && r.c1 <= s.c2)));
                  })
                }
              />
            </Menu>
          }
        />
      </RibbonGroup>
    </>
  );
}
