import { Relationship } from "./common/relationship";
import { PaginatedPage, PaginatedSection } from "./document-pager";
import { RenderedDocumentHandle } from "./html-renderer";
import { Options } from "./options";
import { SerializedPart, SerializedRolePaths } from "./worker/worker-types";
export interface ParseOptions extends Partial<Pick<Options, "breakPages" | "debug" | "ignoreLastRenderedPageBreak" | "inWrapper" | "trimXmlDeclaration">> {
}
export interface RenderOptions extends Partial<Omit<Options, "useWorkerParser" | "workerUrl">> {
}
export interface SnapshotFile {
    path: string;
    buffer: ArrayBuffer;
}
export interface DocxSnapshotMeta {
    version: 1;
    pageCount: number;
    parseOptions: Required<Pick<ParseOptions, "breakPages" | "debug" | "ignoreLastRenderedPageBreak" | "inWrapper" | "trimXmlDeclaration">>;
}
export interface DocxSnapshotPage extends PaginatedPage {
    sections: DocxSnapshotSection[];
}
export interface DocxSnapshotSection extends PaginatedSection {
}
export interface DocxSnapshot {
    meta: DocxSnapshotMeta;
    files: SnapshotFile[];
    rels: Relationship[];
    parts: SerializedPart[];
    rolePaths: SerializedRolePaths;
    pages: DocxSnapshotPage[];
}
export type RenderedSnapshot = RenderedDocumentHandle;
export declare function parseToSnapshot(data: Blob | ArrayBuffer | Uint8Array, options?: ParseOptions): Promise<DocxSnapshot>;
export declare function collectSnapshotTransferables(snapshot: DocxSnapshot): Transferable[];
export declare function renderSnapshot(snapshot: DocxSnapshot, bodyContainer: HTMLElement, styleContainer?: HTMLElement, options?: RenderOptions): Promise<RenderedSnapshot>;
