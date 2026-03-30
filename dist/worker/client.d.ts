import { OpenXmlPackageLoadType, OpenXmlPackageSaveType } from "../common/open-xml-package";
import { Relationship } from "../common/relationship";
import { Options } from "../options";
import { XmlParser } from "../parser/xml-parser";
import { SerializedWordDocument } from "./worker-types";
export interface WorkerParseResult {
    parsed: SerializedWordDocument;
    package: WorkerSessionPackage;
}
export declare function parseDocumentInWorker(data: Blob | ArrayBuffer | Uint8Array, options: Options): Promise<WorkerParseResult>;
declare class ParserWorkerClient {
    private worker;
    private nextRequestId;
    private pending;
    constructor(workerUrl: string);
    parse(buffer: ArrayBuffer, options: Options): Promise<SerializedWordDocument>;
    loadResource(sessionId: string, path: string, outputType: OpenXmlPackageLoadType): Promise<string | ArrayBufferLike | Uint8Array<ArrayBuffer> | Blob>;
    save(sessionId: string, outputType: OpenXmlPackageSaveType): Promise<ArrayBuffer | Uint8Array<ArrayBuffer> | Blob>;
    dispose(sessionId: string): Promise<void>;
    terminate(): void;
    private request;
    private handleMessage;
    private failAll;
}
export declare class WorkerSessionPackage {
    private client;
    private sessionId;
    options: any;
    xmlParser: XmlParser;
    constructor(client: ParserWorkerClient, sessionId: string, options: any);
    get(path: string): boolean;
    update(): void;
    load(path: string, type?: OpenXmlPackageLoadType): Promise<any>;
    save(type?: OpenXmlPackageSaveType): Promise<any>;
    loadRelationships(path?: string): Promise<Relationship[]>;
    parseXmlDocument(text: string): Document;
    dispose(): Promise<void>;
}
export {};
