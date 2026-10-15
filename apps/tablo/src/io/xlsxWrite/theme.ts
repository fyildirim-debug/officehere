/** xl/theme/theme1.xml: kitabın renk ve yazı tipi şemaları, standart biçim şeması. */
import { patchThemeXml, type WorkbookTheme } from "../../model/theme";
import { XML_DECL } from "./xml";

const PH = '<a:schemeClr val="phClr"/>';
const grad = (stops: string) => `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;
const ln = (w: number) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${PH}</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;

const FMT =
  '<a:fmtScheme name="Office">' +
  `<a:fillStyleLst><a:solidFill>${PH}</a:solidFill>` +
  grad('<a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs>') +
  grad('<a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs>') +
  "</a:fillStyleLst>" +
  `<a:lnStyleLst>${ln(6350)}${ln(12700)}${ln(19050)}</a:lnStyleLst>` +
  '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>' +
  `<a:bgFillStyleLst><a:solidFill>${PH}</a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>` +
  grad('<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs>') +
  "</a:bgFillStyleLst></a:fmtScheme>";

/** Kitabın temasıyla tema parçası (renk ve yazı tipi şemaları patchThemeXml ile doldurulur). */
export function themeXml(t: WorkbookTheme): string {
  const base = `${XML_DECL}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name=""></a:clrScheme><a:fontScheme name=""></a:fontScheme>${FMT}</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
  return patchThemeXml(base, t);
}
