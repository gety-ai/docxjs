import { XmlParser } from "../parser/xml-parser";
import { Length } from "./common";
export interface LineSpacing {
    after: Length;
    before: Length;
    line: number;
    lineRule: "atLeast" | "exactly" | "auto";
}
export declare function parseLineSpacing(elem: Element, xml: XmlParser): LineSpacing;
