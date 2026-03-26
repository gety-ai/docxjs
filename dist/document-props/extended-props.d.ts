import { XmlParser } from "../parser/xml-parser";
export interface ExtendedPropsDeclaration {
    template: string;
    totalTime: number;
    pages: number;
    words: number;
    characters: number;
    application: string;
    lines: number;
    paragraphs: number;
    company: string;
    appVersion: string;
}
export declare function parseExtendedProps(root: Element, xmlParser: XmlParser): ExtendedPropsDeclaration;
