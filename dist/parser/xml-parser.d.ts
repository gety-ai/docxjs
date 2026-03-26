import { Length, LengthUsageType } from "../document/common";
export declare function parseXmlString(xmlString: string, trimXmlDeclaration?: boolean): Document;
export declare function serializeXmlString(elem: Node): string;
export declare class XmlParser {
    elements(elem: Element, localName?: string): Element[];
    element(elem: Element, localName: string): Element;
    elementAttr(elem: Element, localName: string, attrLocalName: string): string;
    attrs(elem: Element): any[];
    attr(elem: Element, localName: string): string;
    intAttr(node: Element, attrName: string, defaultValue?: number): number;
    hexAttr(node: Element, attrName: string, defaultValue?: number): number;
    floatAttr(node: Element, attrName: string, defaultValue?: number): number;
    boolAttr(node: Element, attrName: string, defaultValue?: boolean): boolean;
    lengthAttr(node: Element, attrName: string, usage?: LengthUsageType): Length;
}
declare const globalXmlParser: XmlParser;
export default globalXmlParser;
