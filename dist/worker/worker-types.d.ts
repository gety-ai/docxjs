import { Relationship } from "../common/relationship";
export interface SerializedFileEntry {
    path: string;
    buffer: ArrayBufferLike;
}
export interface SerializedPart {
    kind: string;
    path: string;
    rels?: Relationship[];
    [key: string]: any;
}
export interface SerializedWordDocument {
    rels: Relationship[];
    parts: SerializedPart[];
    rolePaths: Record<string, string>;
    files: SerializedFileEntry[];
}
