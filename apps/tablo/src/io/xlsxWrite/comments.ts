/** Notlar: commentsN.xml ve görünümleri (vmlDrawingN.vml). */
import { addr } from "@officehere/formula";
import { esc, XML_DECL } from "./xml";

export interface NoteOut {
  row: number;
  col: number;
  text: string;
}

function runXml(text: string): string {
  return /^\s|\n|\s$/.test(text) ? `<r><t xml:space="preserve">${esc(text)}</t></r>` : `<r><t>${esc(text)}</t></r>`;
}

export function commentsXml(notes: readonly NoteOut[]): string {
  const list = notes.map((n) => `<comment ref="${addr(n.row, n.col)}" authorId="0"><text>${runXml(n.text)}</text></comment>`);
  return `${XML_DECL}<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Author</author></authors><commentList>${list.join("")}</commentList></comments>`;
}

/** Not kutusunun varsayılan yeri: hücrenin sağında, iki sütun genişliğinde, dört satır yüksekliğinde. */
function anchor(row: number, col: number): string {
  const l = col + 1;
  const t = Math.max(row + 1 - 2, 0);
  return [l, 6, t, 14, l + 2, 2, t + 4, 16].join(", ");
}

export function vmlXml(notes: readonly NoteOut[]): string {
  const shapes = notes.map(
    (n, i) =>
      `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" style="position:absolute; margin-left:105.3pt;margin-top:10.5pt;width:97.8pt;height:59.1pt;z-index:1;visibility:hidden" fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">` +
      '<v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/>' +
      '<v:textbox style="mso-direction-alt:auto" inset="1.3mm,1.3mm,2.5mm,2.5mm"><div style="text-align:left"/></v:textbox>' +
      `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${anchor(n.row, n.col)}</x:Anchor><x:Locked>True</x:Locked><x:AutoFill>False</x:AutoFill><x:LockText>True</x:LockText><x:Row>${n.row}</x:Row><x:Column>${n.col}</x:Column></x:ClientData></v:shape>`,
  );
  return (
    XML_DECL +
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
    '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
    shapes.join("") +
    "</xml>"
  );
}
