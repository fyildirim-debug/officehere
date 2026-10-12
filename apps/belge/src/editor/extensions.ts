// Editör şeması ve eklenti listesi.
import { Extension, Node, type AnyExtension } from "@tiptap/core";
import { fontStack } from "@officehere/core";
import StarterKit from "@tiptap/starter-kit";
import { TextStyle, Color, FontFamily, FontSize, BackgroundColor } from "@tiptap/extension-text-style";
import TextAlign from "@tiptap/extension-text-align";
import Highlight from "@tiptap/extension-highlight";
import Subscript from "@tiptap/extension-subscript";
import Superscript from "@tiptap/extension-superscript";
import { ParagraphFormat } from "./paragraphFormat";
import { CharStyle, CommentMark, Deletion, FormatChange, Insertion, TextStyleExtras, Underline } from "./marks";
import { Bookmark, ListStyles, PageBreak, PageField, TableOfContents } from "./nodes";
import { BelgeTable, BelgeTableCell, BelgeTableHeader, BelgeTableRow } from "./table";
import { BelgeImage } from "./image";
import { Pagination, type PaginationController } from "./pagination";
import { Search } from "./search";
import { ShowMarks } from "./showMarks";
import { TrackChanges } from "./trackChanges";
import { createWordKeymap, type EditorHooks } from "./keymap";
import { Note, Notes } from "./notes";
import { PasteRules } from "./paste";
import { TextBox } from "./textbox";
import { BelgeDocument, ColumnBreak, Section, SectionKeys, createSectionLayout } from "./sections";
import { AutoCorrect } from "./autocorrect";
import { TableHandles } from "./tableHandles";
import { TabStopsAttr } from "./tabStops";

/** Yazı tipi adı yedek zinciriyle (yüklü olmayan yazı tiplerinde ölçü uyumlu açık yazı tipine düşer). */
const BelgeFontFamily = FontFamily.extend({
  addGlobalAttributes() {
    return [
      {
        types: ["textStyle"],
        attributes: {
          fontFamily: {
            default: null,
            parseHTML: (el: HTMLElement) => {
              const v = el.style.fontFamily;
              if (!v) return null;
              return v.split(",")[0].replace(/["']/g, "").trim() || null;
            },
            renderHTML: (a: Record<string, unknown>) => (a.fontFamily ? { style: `font-family: ${fontStack(String(a.fontFamily))}` } : {}),
          },
        },
      },
    ];
  },
});

const Align = TextAlign.extend({
  // Hizalama kısayolları Word'deki gibi Ctrl+L/E/R/J (wordKeymap'te).
  addKeyboardShortcuts() {
    return {};
  },
}).configure({ types: ["heading", "paragraph"], alignments: ["left", "center", "right", "justify"] });

/** Ortak içerik şeması (belge kökü hariç). */
function contentExtensions(): AnyExtension[] {
  return [
    StarterKit.configure({
      document: false,
      blockquote: false,
      code: false,
      codeBlock: false,
      underline: false,
      trailingNode: false,
      heading: { levels: [1, 2, 3, 4, 5, 6] },
      link: { openOnClick: false, autolink: true, linkOnPaste: true, defaultProtocol: "https", HTMLAttributes: { rel: null, target: null, class: "bl-link" } },
      undoRedo: { depth: 300, newGroupDelay: 600 },
      dropcursor: { color: "#000000", width: 2 },
    }),
    TextStyle,
    TextStyleExtras,
    Color,
    BelgeFontFamily,
    FontSize,
    BackgroundColor,
    Align,
    Highlight.configure({ multicolor: true }),
    Subscript,
    Superscript,
    Underline,
    CharStyle,
    CommentMark,
    Insertion,
    Deletion,
    FormatChange,
    ParagraphFormat,
    TabStopsAttr,
    ListStyles,
    PageBreak,
    TableOfContents,
    PageField,
    Bookmark,
    BelgeTable,
    BelgeTableRow,
    BelgeTableHeader,
    BelgeTableCell,
    BelgeImage,
    Note,
    TextBox,
  ];
}

/** Belge içeriği şeması: belge > bölüm+ > blok+ (statik dönüşümler ve testler için de kullanılır). */
export function schemaExtensions(): AnyExtension[] {
  return [...contentExtensions(), BelgeDocument, Section, ColumnBreak];
}

/** Üst/alt bilgi kökü: bölümsüz blok dizisi. */
const PlainDocument = Node.create({ name: "doc", topNode: true, content: "block+" });

/** Üst/alt bilgi ve notlarda sayfa/sütun sonu eklenemez (Word'deki gibi Ctrl+Enter etkisizdir). */
const NoBreakKeys = Extension.create({
  name: "belgeNoBreakKeys",
  priority: 1200,
  addKeyboardShortcuts() {
    return { "Mod-Enter": () => true, "Mod-Shift-Enter": () => true };
  },
});

export function editorExtensions(controller: PaginationController, hooks: EditorHooks): AnyExtension[] {
  return [
    ...schemaExtensions(),
    Pagination.configure({ controller }),
    Notes,
    Search,
    ShowMarks,
    TrackChanges,
    createWordKeymap(hooks),
    SectionKeys,
    PasteRules,
    AutoCorrect,
    TableHandles,
    createSectionLayout({ enabled: () => controller.geometry.enabled, defaultPage: () => controller.geometry.defaultPage, tiled: () => controller.geometry.tiles > 0 }),
  ];
}

/** Üst/alt bilgi editörü: gövdeyle aynı biçimlendirme, sayfalama yok. */
export function headerFooterExtensions(): AnyExtension[] {
  return [...contentExtensions(), PlainDocument, NoBreakKeys, AutoCorrect];
}
