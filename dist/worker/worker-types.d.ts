import { Relationship } from "../common/relationship";
import { OpenXmlPackageLoadType, OpenXmlPackageSaveType } from "../common/open-xml-package";
export interface SerializedPart {
    kind: string;
    path: string;
    rels?: Relationship[];
    [key: string]: any;
}
export type SerializedRolePaths = Record<string, string>;
export interface SerializedWordDocument {
    sessionId: string;
    rels: Relationship[];
    parts: SerializedPart[];
    rolePaths: SerializedRolePaths;
}
export interface WorkerRequestBase {
    requestId: number;
}
export interface ParseDocumentRequest extends WorkerRequestBase {
    type: "parse";
    buffer: ArrayBuffer;
    options: any;
}
export interface LoadResourceRequest extends WorkerRequestBase {
    type: "load-resource";
    sessionId: string;
    path: string;
    outputType: OpenXmlPackageLoadType;
}
export interface SaveDocumentRequest extends WorkerRequestBase {
    type: "save";
    sessionId: string;
    outputType: OpenXmlPackageSaveType;
}
export interface DisposeSessionRequest extends WorkerRequestBase {
    type: "dispose";
    sessionId: string;
}
export type WorkerRequest = ParseDocumentRequest | LoadResourceRequest | SaveDocumentRequest | DisposeSessionRequest;
export interface ParsedDocumentResponse extends WorkerRequestBase {
    type: "parsed";
    payload: SerializedWordDocument;
}
export interface ResourceResponse extends WorkerRequestBase {
    type: "resource";
    outputType: OpenXmlPackageLoadType;
    value: string | ArrayBufferLike;
}
export interface SaveResponse extends WorkerRequestBase {
    type: "saved";
    outputType: OpenXmlPackageSaveType;
    value: ArrayBufferLike;
}
export interface DisposedResponse extends WorkerRequestBase {
    type: "disposed";
}
export interface NullResponse extends WorkerRequestBase {
    type: "null";
}
export interface ErrorResponse extends WorkerRequestBase {
    type: "error";
    error: string;
}
export type WorkerResponse = ParsedDocumentResponse | ResourceResponse | SaveResponse | DisposedResponse | NullResponse | ErrorResponse;
