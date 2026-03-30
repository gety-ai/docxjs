import { OpenXmlPackageLoadType, OpenXmlPackageSaveType } from "../common/open-xml-package";
import { parseRelationships, Relationship } from "../common/relationship";
import { Options } from "../options";
import { parseXmlString, XmlParser } from "../parser/xml-parser";
import { splitPath } from "../utils";
import {
    ErrorResponse,
    ParsedDocumentResponse,
    ResourceResponse,
    SaveResponse,
    SerializedWordDocument,
    WorkerResponse
} from "./worker-types";

export interface WorkerParseResult {
    parsed: SerializedWordDocument;
    package: WorkerSessionPackage;
}

export async function parseDocumentInWorker(data: Blob | ArrayBuffer | Uint8Array, options: Options): Promise<WorkerParseResult> {
    const workerUrl = resolveWorkerUrl(options.workerUrl);

    if (!workerUrl || typeof Worker === "undefined") {
        return null;
    }

    const client = new ParserWorkerClient(workerUrl);

    try {
        const buffer = await toArrayBuffer(data);
        const parsed = await client.parse(buffer, options);

        return {
            parsed,
            package: new WorkerSessionPackage(client, parsed.sessionId, options)
        };
    } catch (error) {
        client.terminate();
        throw error;
    }
}

class ParserWorkerClient {
    private worker: Worker;
    private nextRequestId = 1;
    private pending = new Map<number, {
        resolve: (value: WorkerResponse) => void;
        reject: (reason?: any) => void;
    }>();

    constructor(workerUrl: string) {
        this.worker = new Worker(workerUrl);
        this.worker.onmessage = event => this.handleMessage(event.data as WorkerResponse);
        this.worker.onerror = event => this.failAll(event.error ?? new Error(event.message));
    }

    async parse(buffer: ArrayBuffer, options: Options): Promise<SerializedWordDocument> {
        const response = await this.request<ParsedDocumentResponse>({
            type: "parse",
            buffer,
            options: prepareWorkerOptions(options)
        }, [buffer]);

        return response.payload;
    }

    async loadResource(sessionId: string, path: string, outputType: OpenXmlPackageLoadType) {
        const response = await this.request<ResourceResponse | WorkerResponse>({
            type: "load-resource",
            sessionId,
            path,
            outputType
        });

        if (response.type === "null")
            return null;

        if (outputType === "string")
            return (response as ResourceResponse).value;

        const buffer = (response as ResourceResponse).value as ArrayBuffer;

        switch (outputType) {
            case "uint8array":
                return new Uint8Array(buffer);
            case "blob":
                return new Blob([new Uint8Array(buffer)]);
            case "arraybuffer":
            default:
                return buffer;
        }
    }

    async save(sessionId: string, outputType: OpenXmlPackageSaveType) {
        const response = await this.request<SaveResponse>({
            type: "save",
            sessionId,
            outputType
        });

        const buffer = response.value as ArrayBuffer;

        switch (outputType) {
            case "uint8array":
                return new Uint8Array(buffer);
            case "blob":
                return new Blob([new Uint8Array(buffer)]);
            case "arraybuffer":
            default:
                return buffer;
        }
    }

    async dispose(sessionId: string) {
        if (!this.worker)
            return;

        try {
            await this.request({
                type: "dispose",
                sessionId
            });
        } finally {
            this.terminate();
        }
    }

    terminate() {
        if (!this.worker)
            return;

        this.worker.terminate();
        this.worker = null;
        this.failAll(new Error("Parser worker terminated"));
    }

    private request<T extends WorkerResponse>(message: Record<string, any>, transfer: Transferable[] = []): Promise<T> {
        const requestId = this.nextRequestId++;

        return new Promise((resolve, reject) => {
            this.pending.set(requestId, {
                resolve: resolve as any,
                reject
            });

            this.worker.postMessage({
                ...message,
                requestId
            }, transfer);
        });
    }

    private handleMessage(message: WorkerResponse) {
        const pending = this.pending.get(message.requestId);

        if (!pending)
            return;

        this.pending.delete(message.requestId);

        if (message.type === "error") {
            pending.reject(new Error((message as ErrorResponse).error ?? "Unknown parser worker error"));
        } else {
            pending.resolve(message);
        }
    }

    private failAll(error: Error) {
        if (this.pending.size === 0)
            return;

        const pending = Array.from(this.pending.values());
        this.pending.clear();

        for (const entry of pending) {
            entry.reject(error);
        }
    }
}

export class WorkerSessionPackage {
    xmlParser: XmlParser = new XmlParser();

    constructor(
        private client: ParserWorkerClient,
        private sessionId: string,
        public options: any
    ) {
    }

    get(path: string) {
        return path ? true : null;
    }

    update() {
        throw new Error("DOCX: update() is not supported for worker-backed packages");
    }

    load(path: string, type: OpenXmlPackageLoadType = "string"): Promise<any> {
        return this.client.loadResource(this.sessionId, normalizePath(path), type);
    }

    save(type: OpenXmlPackageSaveType = "blob"): Promise<any> {
        return this.client.save(this.sessionId, type);
    }

    async loadRelationships(path: string = null): Promise<Relationship[]> {
        let relsPath = `_rels/.rels`;

        if (path != null) {
            const [folder, fileName] = splitPath(path);
            relsPath = `${folder}_rels/${fileName}.rels`;
        }

        const text = await this.load(relsPath, "string");
        return text ? parseRelationships(this.parseXmlDocument(text).firstElementChild, this.xmlParser) : null;
    }

    parseXmlDocument(text: string): Document {
        return parseXmlString(text, this.options.trimXmlDeclaration);
    }

    dispose() {
        return this.client.dispose(this.sessionId);
    }
}

function prepareWorkerOptions(options: Options) {
    return JSON.parse(JSON.stringify({
        ...options,
        workerUrl: undefined
    }));
}

function resolveWorkerUrl(explicitUrl?: string | URL) {
    if (explicitUrl)
        return explicitUrl instanceof URL ? explicitUrl.href : explicitUrl.toString();

    if (typeof document === "undefined")
        return null;

    const scripts = Array.from(document.scripts ?? []).reverse();

    for (const script of scripts) {
        const src = script.src;

        if (!src)
            continue;

        if (/docx-preview(?:\.min)?\.js(?:\?.*)?$/i.test(src)) {
            return src.replace(/docx-preview(?:\.min)?\.js(?:\?.*)?$/i, "docx-preview-worker.js");
        }
    }

    return null;
}

async function toArrayBuffer(data: Blob | ArrayBuffer | Uint8Array): Promise<ArrayBuffer> {
    if (data instanceof ArrayBuffer)
        return data;

    if (data instanceof Uint8Array)
        return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;

    if (data && typeof (data as Blob).arrayBuffer === "function")
        return await (data as Blob).arrayBuffer();

    throw new Error("Unsupported input type for parser worker");
}

function normalizePath(path: string) {
    return path.startsWith("/") ? path.substring(1) : path;
}
