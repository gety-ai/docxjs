import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { WmlBaseNote, WmlEndnote, WmlFootnote } from "./elements";
export declare class BaseNotePart<T extends WmlBaseNote> extends Part {
    protected _documentParser: DocumentParser;
    notes: T[];
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
}
export declare class FootnotesPart extends BaseNotePart<WmlFootnote> {
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
    parseXml(root: Element): void;
}
export declare class EndnotesPart extends BaseNotePart<WmlEndnote> {
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
    parseXml(root: Element): void;
}
