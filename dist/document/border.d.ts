import { XmlParser } from "../parser/xml-parser";
import { Length } from "./common";
export interface Border {
    color: string;
    type: string;
    size: Length;
    frame: boolean;
    shadow: boolean;
    offset: Length;
}
export interface Borders {
    top: Border;
    left: Border;
    right: Border;
    bottom: Border;
}
export declare function parseBorder(elem: Element, xml: XmlParser): Border;
export declare function parseBorders(elem: Element, xml: XmlParser): Borders;
