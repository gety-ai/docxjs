import { XmlParser } from "../parser/xml-parser";
import { CommonProperties } from "./common";
import { OpenXmlElement } from "./dom";
export interface WmlRun extends OpenXmlElement, RunProperties {
    id?: string;
    verticalAlign?: string;
    fieldRun?: boolean;
}
export interface RunProperties extends CommonProperties {
}
export declare function parseRunProperties(elem: Element, xml: XmlParser): RunProperties;
export declare function parseRunProperty(elem: Element, props: RunProperties, xml: XmlParser): boolean;
