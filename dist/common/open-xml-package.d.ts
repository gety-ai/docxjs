import { XmlParser } from "../parser/xml-parser";
import { Relationship } from "./relationship";
export interface OpenXmlPackageOptions {
    trimXmlDeclaration: boolean;
    keepOrigin: boolean;
}
export type OpenXmlPackageLoadType = "string" | "uint8array" | "arraybuffer" | "blob";
export type OpenXmlPackageSaveType = "blob" | "uint8array" | "arraybuffer";
export declare class OpenXmlPackage {
    private _files;
    options: OpenXmlPackageOptions;
    xmlParser: XmlParser;
    private decoder;
    private encoder;
    constructor(_files: Record<string, Uint8Array>, options: OpenXmlPackageOptions);
    get(path: string): any;
    update(path: string, content: any): void;
    static load(input: Blob | any, options: OpenXmlPackageOptions): Promise<OpenXmlPackage>;
    static fromFiles(files: Record<string, Uint8Array>, options: OpenXmlPackageOptions): OpenXmlPackage;
    save(type?: OpenXmlPackageSaveType): Promise<any>;
    load(path: string, type?: OpenXmlPackageLoadType): Promise<any>;
    loadRelationships(path?: string): Promise<Relationship[]>;
    parseXmlDocument(txt: string): Document;
}
