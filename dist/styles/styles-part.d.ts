import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { IDomStyle } from "../document/style";
export declare class StylesPart extends Part {
    styles: IDomStyle[];
    private _documentParser;
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
    parseXml(root: Element): void;
}
