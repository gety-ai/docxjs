import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { OpenXmlElement } from "../document/dom";
import { WmlHeader, WmlFooter } from "./elements";
export declare abstract class BaseHeaderFooterPart<T extends OpenXmlElement = OpenXmlElement> extends Part {
    rootElement: T;
    private _documentParser;
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
    parseXml(root: Element): void;
    protected abstract createRootElement(): T;
}
export declare class HeaderPart extends BaseHeaderFooterPart<WmlHeader> {
    protected createRootElement(): WmlHeader;
}
export declare class FooterPart extends BaseHeaderFooterPart<WmlFooter> {
    protected createRootElement(): WmlFooter;
}
