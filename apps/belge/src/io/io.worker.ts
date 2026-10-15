// Belge dosya G/Ç Worker'ı: açma (bayt → belge modeli), .docx/.odt/.rtf/.md yazma ve otomatik kayıt doğrulaması.
// Worker'da DOMParser yoktur: XML yerleşik ayrıştırıcıyla xmldom DOM'unda kurulur (xmlDom.ts).
import { exposeIo } from "@officehere/core";
import { installXmlDom } from "./xmlDom";
import { handlers } from "./ioHandlers";

installXmlDom();
exposeIo(handlers);
