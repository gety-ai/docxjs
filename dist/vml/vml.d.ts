import { DocumentParser } from '../document-parser';
import { OpenXmlElementBase, DomType } from '../document/dom';
export declare class VmlElement extends OpenXmlElementBase {
    type: DomType;
    tagName: string;
    cssStyleText?: string;
    attrs: Record<string, string>;
    wrapType?: string;
    imageHref?: {
        id: string;
        title: string;
    };
}
export declare function parseVmlElement(elem: Element, parser: DocumentParser): VmlElement;
