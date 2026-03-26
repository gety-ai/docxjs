import { XmlParser } from "../parser/xml-parser";
import { OpenXmlElement } from "./dom";
export interface WmlBookmarkStart extends OpenXmlElement {
    id: string;
    name: string;
    colFirst: number;
    colLast: number;
}
export interface WmlBookmarkEnd extends OpenXmlElement {
    id: string;
}
export declare function parseBookmarkStart(elem: Element, xml: XmlParser): WmlBookmarkStart;
export declare function parseBookmarkEnd(elem: Element, xml: XmlParser): WmlBookmarkEnd;
