/** Yardım sekmesi: kısayollar, hakkında, işlev başvurusu. */
import { Info24Regular, Keyboard24Regular, MathFormula24Regular } from "@fluentui/react-icons";
import { ReportBugButton, RibbonButton, RibbonGroup } from "@officehere/ui";
import type { Controller } from "../controller";

export function HelpTab({ ctl }: { ctl: Controller }) {
  return (
    <RibbonGroup label="Yardım" icon={<Info24Regular />}>
      <RibbonButton size="large" icon={<Keyboard24Regular />} label="Klavye Kısayolları" description="Tablo'daki klavye kısayollarının listesini gösterir." onClick={() => ctl.openDialog("shortcuts")} />
      <RibbonButton size="large" icon={<MathFormula24Regular />} label="İşlev Başvurusu" description="Tüm işlevleri kategorilere göre arayın ve açıklamalarını okuyun." onClick={() => ctl.openDialog("insertFunction")} />
      <RibbonButton size="large" icon={<Info24Regular />} label="Hakkında" description="Sürüm ve lisans bilgileri." onClick={() => ctl.openDialog("about")} />
      <ReportBugButton app="tablo" />
    </RibbonGroup>
  );
}
