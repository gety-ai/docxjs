import { Part } from "../common/part";
import { OpenXmlPackage } from "../common/open-xml-package";
export type CommentsExtended = {
    paraId: string;
    paraIdParent?: string;
    done: boolean;
};
export declare class CommentsExtendedPart extends Part {
    comments: CommentsExtended[];
    commentMap: Record<string, CommentsExtended>;
    constructor(pkg: OpenXmlPackage, path: string);
    parseXml(root: Element): void;
}
