import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { DocumentElement } from "./document";
export declare class DocumentPart extends Part {
    private _documentParser;
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
    body: DocumentElement;
    parseXml(root: Element): void;
}
