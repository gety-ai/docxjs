import { OpenXmlPackage } from "../common/open-xml-package";
import { Part } from "../common/part";
import { DocumentParser } from "../document-parser";
import { IDomNumbering } from "../document/dom";
import { AbstractNumbering, Numbering, NumberingBulletPicture, NumberingPartProperties } from "./numbering";
export declare class NumberingPart extends Part implements NumberingPartProperties {
    private _documentParser;
    constructor(pkg: OpenXmlPackage, path: string, parser: DocumentParser);
    numberings: Numbering[];
    abstractNumberings: AbstractNumbering[];
    bulletPictures: NumberingBulletPicture[];
    domNumberings: IDomNumbering[];
    parseXml(root: Element): void;
}
