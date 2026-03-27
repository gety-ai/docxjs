import { Options } from "../docx-preview";
import { SerializedWordDocument } from "./worker-types";

export async function parseDocumentInWorker(data: Blob | ArrayBuffer | Uint8Array, options: Options): Promise<SerializedWordDocument> {
    const workerUrl = resolveWorkerUrl(options.workerUrl);

    if (!workerUrl || typeof Worker === "undefined") {
        return null;
    }

    const buffer = await toArrayBuffer(data);

    return new Promise((resolve, reject) => {
        const worker = new Worker(workerUrl);

        worker.onmessage = event => {
            const payload = event.data;
            worker.terminate();

            if (payload?.type === "parsed") {
                resolve(payload.payload as SerializedWordDocument);
            } else {
                reject(new Error(payload?.error ?? "Unknown parser worker error"));
            }
        };

        worker.onerror = event => {
            worker.terminate();
            reject(event.error ?? new Error(event.message));
        };

        worker.postMessage({
            type: "parse",
            buffer,
            options: JSON.parse(JSON.stringify(options))
        }, [buffer]);
    });
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

async function toArrayBuffer(data: Blob | ArrayBuffer | Uint8Array) {
    if (data instanceof ArrayBuffer)
        return data;

    if (data instanceof Uint8Array)
        return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);

    if (data && typeof (data as Blob).arrayBuffer === "function")
        return await (data as Blob).arrayBuffer();

    throw new Error("Unsupported input type for parser worker");
}
