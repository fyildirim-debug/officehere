// ODF geçiş katmanı (2. aşama P5.5, teknik tasarım K§1.2.5): .odt/.ods/.odp'nin özgün paketi kaynak kaydında
// (@officehere/core/opc; ODF ZIP'i OPC değildir, yalnız ZIP okuyucu ve kayıt ortaktır). Okurken manifest listesi
// ve opak çizim nesneleri (draw:frame + andığı otomatik stiller), yazarken OdfCarry: settings.xml/manifest.rdf
// taşıma, Thumbnails/ düşürme, Object N/ alt belgeleri ve yedek görüntüleri ham kopya, manifest yeniden kurulumu.
export * from "./manifest";
export * from "./xml";
export * from "./carry";
export * from "./read";
