import { XmlParser } from "../parser/xml-parser";
export interface CorePropsDeclaration {
    title: string;
    description: string;
    subject: string;
    creator: string;
    keywords: string;
    language: string;
    lastModifiedBy: string;
    revision: number;
}
export declare function parseCoreProps(root: Element, xmlParser: XmlParser): CorePropsDeclaration;
