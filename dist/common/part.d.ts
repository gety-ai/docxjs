import { OpenXmlPackage } from "./open-xml-package";
import { Relationship } from "./relationship";
export declare class Part {
    protected _package: OpenXmlPackage;
    path: string;
    protected _xmlDocument: Document;
    rels: Relationship[];
    constructor(_package: OpenXmlPackage, path: string);
    load(): Promise<any>;
    save(): void;
    protected parseXml(root: Element): void;
}
