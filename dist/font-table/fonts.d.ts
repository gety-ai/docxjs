import { XmlParser } from "../parser/xml-parser";
export interface FontDeclaration {
    name: string;
    altName: string;
    family: string;
    embedFontRefs: EmbedFontRef[];
}
export interface EmbedFontRef {
    id: string;
    key: string;
    type: 'regular' | 'bold' | 'italic' | 'boldItalic';
}
export declare function parseFonts(root: Element, xml: XmlParser): FontDeclaration[];
export declare function parseFont(elem: Element, xml: XmlParser): FontDeclaration;
export declare function parseEmbedFontRef(elem: Element, xml: XmlParser): EmbedFontRef;
