import { unzipSync, zipSync } from "fflate";
import { parseXmlString, XmlParser } from "../parser/xml-parser";
import { splitPath } from "../utils";
import { parseRelationships, Relationship } from "./relationship";

export interface OpenXmlPackageOptions {
    trimXmlDeclaration: boolean,
    keepOrigin: boolean,
}

export type OpenXmlPackageLoadType = "string" | "uint8array" | "arraybuffer" | "blob";
export type OpenXmlPackageSaveType = "blob" | "uint8array" | "arraybuffer";

export class OpenXmlPackage {
    xmlParser: XmlParser = new XmlParser();
    private decoder = new TextDecoder();
    private encoder = new TextEncoder();

    constructor(private _files: Record<string, Uint8Array>, public options: OpenXmlPackageOptions) {
    }

    get(path: string): any {
        const p = normalizePath(path);
        return this._files[p] ?? this._files[p.replace(/\//g, "\\")] ?? null;
    }

    update(path: string, content: any) {
        this._files[normalizePath(path)] = toUint8Array(content, this.encoder);
    }

    static async load(input: Blob | any, options: OpenXmlPackageOptions): Promise<OpenXmlPackage> {
        const data = await inputToUint8Array(input);
        return new OpenXmlPackage(normalizeFiles(unzipSync(data)), options);
    }

    static fromFiles(files: Record<string, Uint8Array>, options: OpenXmlPackageOptions): OpenXmlPackage {
        return new OpenXmlPackage(normalizeFiles(files), options);
    }

    save(type: OpenXmlPackageSaveType = "blob"): Promise<any>  {
        const zipped = zipSync(this._files);

        switch (type) {
            case "uint8array":
                return Promise.resolve(zipped);
            case "arraybuffer":
                return Promise.resolve(zipped.buffer.slice(zipped.byteOffset, zipped.byteOffset + zipped.byteLength));
            case "blob":
            default:
                return Promise.resolve(new Blob([new Uint8Array(zipped)]));
        }
    }

    load(path: string, type: OpenXmlPackageLoadType = "string"): Promise<any> {
        const file = this.get(path);
        if (!file)
            return Promise.resolve(null);

        switch (type) {
            case "uint8array":
                return Promise.resolve(file.slice());
            case "arraybuffer":
                return Promise.resolve(file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength));
            case "blob":
                return Promise.resolve(new Blob([new Uint8Array(file)]));
            case "string":
            default:
                return Promise.resolve(this.decoder.decode(file));
        }
    }

    async loadRelationships(path: string = null): Promise<Relationship[]> {
        let relsPath = `_rels/.rels`;

        if (path != null) {
            const [f, fn] = splitPath(path);
            relsPath = `${f}_rels/${fn}.rels`;
        }

        const txt = await this.load(relsPath);
		return txt ? parseRelationships(this.parseXmlDocument(txt).firstElementChild, this.xmlParser) : null;
    }

    /** @internal */
    parseXmlDocument(txt: string): Document {
        return parseXmlString(txt, this.options.trimXmlDeclaration);
    }
}

function normalizePath(path: string) {
    return path.startsWith('/') ? path.substr(1) : path;
}

function normalizeFiles(files: Record<string, Uint8Array>) {
    const result: Record<string, Uint8Array> = {};

    for (const [path, file] of Object.entries(files ?? {})) {
        result[normalizePath(path)] = file;
    }

    return result;
}

async function inputToUint8Array(input: Blob | ArrayBuffer | Uint8Array | ArrayBufferView | any) {
    if (input instanceof Uint8Array)
        return input.slice();

    if (input instanceof ArrayBuffer)
        return new Uint8Array(input);

    if (ArrayBuffer.isView(input))
        return new Uint8Array(input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength));

    if (input instanceof Blob)
        return new Uint8Array(await input.arrayBuffer());

    if (typeof input?.arrayBuffer === "function")
        return new Uint8Array(await input.arrayBuffer());

    throw new Error("Unsupported input type for OpenXmlPackage.load");
}

function toUint8Array(value: any, encoder: TextEncoder) {
    if (value instanceof Uint8Array)
        return value;

    if (value instanceof ArrayBuffer)
        return new Uint8Array(value);

    if (typeof value === "string")
        return encoder.encode(value);

    if (value instanceof Blob)
        throw new Error("Blob updates are not supported for in-memory OpenXmlPackage");

    return new Uint8Array(value);
}
