import { XmlParser } from "../parser/xml-parser";
export interface CustomProperty {
    formatId: string;
    name: string;
    type: string;
    value: string;
}
export declare function parseCustomProps(root: Element, xml: XmlParser): CustomProperty[];
