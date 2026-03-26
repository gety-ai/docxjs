import { Part } from "../common/part";
import { OpenXmlPackage } from "../common/open-xml-package";
import { DocumentParser } from "../document-parser";
import { WmlComment } from "./elements";
export declare class CommentsPart extends Part {
    protected _documentParser: DocumentParser;
    comments: WmlComment[];
    commentMap: Record<string, WmlComment>;
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
    parseXml(root: Element): void;
}
