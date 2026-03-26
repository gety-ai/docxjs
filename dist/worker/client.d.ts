import { Options } from "../docx-preview";
import { SerializedWordDocument } from "./worker-types";
export declare function parseDocumentInWorker(data: Blob | ArrayBuffer | Uint8Array, options: Options): Promise<SerializedWordDocument>;
