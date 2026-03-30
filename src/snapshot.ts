import { unzipSync } from "fflate";

import { Relationship, RelationshipTypes, parseRelationships } from "./common/relationship";
import { OpenXmlPackageLoadType } from "./common/open-xml-package";
import { Part } from "./common/part";
import { CommentsExtendedPart } from "./comments/comments-extended-part";
import { CommentsPart } from "./comments/comments-part";
import { CorePropsPart } from "./document-props/core-props-part";
import { CustomPropsPart } from "./document-props/custom-props-part";
import { ExtendedPropsPart } from "./document-props/extended-props-part";
import { DocumentPager, DocumentPagingOptions, PaginatedPage, PaginatedSection } from "./document-pager";
import { DocumentPart } from "./document/document-part";
import { DocumentParser } from "./document-parser";
import { FontTablePart } from "./font-table/font-table";
import { FooterPart, HeaderPart } from "./header-footer/parts";
import { HtmlRenderer, RenderedDocumentHandle } from "./html-renderer";
import { EndnotesPart, FootnotesPart } from "./notes/parts";
import { NumberingPart } from "./numbering/numbering-part";
import { defaultOptions, Options } from "./options";
import { XmlParser } from "./parser/xml-parser";
import { parseXmlStringWithTxml } from "./parser/txml-parser";
import { SettingsPart } from "./settings/settings-part";
import { StylesPart } from "./styles/styles-part";
import { ThemePart } from "./theme/theme-part";
import { resolvePath, splitPath } from "./utils";
import { WordDocument } from "./word-document";
import { SerializedPart, SerializedRolePaths } from "./worker/worker-types";

const topLevelRels = [
    { type: RelationshipTypes.OfficeDocument, target: "word/document.xml" },
    { type: RelationshipTypes.ExtendedProperties, target: "docProps/app.xml" },
    { type: RelationshipTypes.CoreProperties, target: "docProps/core.xml" },
    { type: RelationshipTypes.CustomProperties, target: "docProps/custom.xml" },
];

type PartKind =
    | "document"
    | "fontTable"
    | "numbering"
    | "styles"
    | "theme"
    | "footnotes"
    | "endnotes"
    | "footer"
    | "header"
    | "coreProps"
    | "extendedProps"
    | "customProps"
    | "settings"
    | "comments"
    | "commentsExtended";

export interface ParseOptions extends Partial<Pick<Options,
    "breakPages" |
    "debug" |
    "ignoreLastRenderedPageBreak" |
    "inWrapper" |
    "trimXmlDeclaration"
>> {
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
    parseOptions: Required<Pick<ParseOptions,
        "breakPages" |
        "debug" |
        "ignoreLastRenderedPageBreak" |
        "inWrapper" |
        "trimXmlDeclaration"
    >>;
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

export async function parseToSnapshot(
    data: Blob | ArrayBuffer | Uint8Array,
    options?: ParseOptions,
): Promise<DocxSnapshot> {
    const normalized = normalizeParseOptions(options);
    const builder = new SnapshotBuilder(normalized);
    return await builder.build(data);
}

export function collectSnapshotTransferables(
    snapshot: DocxSnapshot,
): Transferable[] {
    return (snapshot.files ?? []).map(file => file.buffer as Transferable);
}

export async function renderSnapshot(
    snapshot: DocxSnapshot,
    bodyContainer: HTMLElement,
    styleContainer?: HTMLElement,
    options?: RenderOptions,
): Promise<RenderedSnapshot> {
    validateSnapshot(snapshot);

    const normalized = normalizeRenderOptions(options);
    validateRenderOptions(snapshot, normalized);

    const document = WordDocument.fromSnapshot(snapshot, normalized);
    const renderer = new HtmlRenderer(bodyContainer.ownerDocument ?? window.document);
    return await renderer.render(document, bodyContainer, styleContainer, normalized);
}

class SnapshotOpenXmlPackage {
    xmlParser = new XmlParser();
    decoder = new TextDecoder();

    constructor(public files: Record<string, Uint8Array>, public options: Pick<Options, "trimXmlDeclaration">) {
    }

    get(path: string) {
        const normalized = normalizePath(path);
        return this.files[normalized] ?? this.files[normalized.replace(/\//g, "\\")] ?? null;
    }

    async load(path: string, type: OpenXmlPackageLoadType = "string"): Promise<any> {
        const file = this.get(path);

        if (!file)
            return null;

        switch (type) {
            case "uint8array":
                return file.slice();
            case "arraybuffer":
                return toArrayBuffer(file);
            case "blob":
                return new Blob([file.slice()]);
            case "string":
            default:
                return this.decoder.decode(file);
        }
    }

    async loadRelationships(path: string = null): Promise<Relationship[]> {
        let relsPath = `_rels/.rels`;

        if (path != null) {
            const [folder, fileName] = splitPath(path);
            relsPath = `${folder}_rels/${fileName}.rels`;
        }

        const text = await this.load(relsPath, "string");
        return text ? parseRelationships(this.parseXmlDocument(text).firstElementChild as any, this.xmlParser) : null;
    }

    parseXmlDocument(text: string): any {
        return parseXmlStringWithTxml(text, this.options.trimXmlDeclaration);
    }
}

class SnapshotBuilder {
    private parser: DocumentParser;
    private package: SnapshotOpenXmlPackage;
    private rels: Relationship[] = [];
    private parts: Part[] = [];
    private partsMap: Record<string, Part> = {};
    private rolePaths: SerializedRolePaths = {};
    private documentPart: DocumentPart = null;
    private stylesPart: StylesPart = null;

    constructor(private options: Required<Pick<ParseOptions,
        "breakPages" |
        "debug" |
        "ignoreLastRenderedPageBreak" |
        "inWrapper" |
        "trimXmlDeclaration"
    >>) {
        this.parser = new DocumentParser(options);
    }

    async build(data: Blob | ArrayBuffer | Uint8Array): Promise<DocxSnapshot> {
        const files = unzipSync(await inputToUint8Array(data));
        this.package = new SnapshotOpenXmlPackage(normalizeFiles(files), this.options);
        this.rels = await this.package.loadRelationships();

        await Promise.all(topLevelRels.map(rel => {
            const relationship = this.rels.find(x => x.type === rel.type) ?? rel;
            return this.loadRelationshipPart(relationship.target, relationship.type);
        }));

        const pagerOptions: DocumentPagingOptions = {
            breakPages: this.options.breakPages,
            className: defaultOptions.className,
            debug: this.options.debug,
            ignoreLastRenderedPageBreak: this.options.ignoreLastRenderedPageBreak,
            inWrapper: this.options.inWrapper
        };
        const styleMap = this.stylesPart?.styles
            ? DocumentPager.createStyleMap(cloneSerializable(this.stylesPart.styles), pagerOptions)
            : {};
        const pages = new DocumentPager(pagerOptions, styleMap).buildPages(this.documentPart.body) as DocxSnapshotPage[];

        return {
            meta: {
                version: 1,
                pageCount: pages.length,
                parseOptions: this.options
            },
            files: this.collectSnapshotFiles(),
            rels: this.rels,
            parts: this.parts.map(part => this.serializePart(part)).filter(Boolean),
            rolePaths: this.rolePaths,
            pages
        };
    }

    private async loadRelationshipPart(path: string, type: string): Promise<Part> {
        const normalizedPath = normalizePath(path);

        if (this.partsMap[normalizedPath])
            return this.partsMap[normalizedPath];

        if (!this.package.get(normalizedPath))
            return null;

        let part: Part = null;
        let partKind: PartKind = null;

        switch (type) {
            case RelationshipTypes.OfficeDocument:
                this.documentPart = part = new DocumentPart(this.package as any, normalizedPath, this.parser);
                partKind = "document";
                this.rolePaths.documentPart = normalizedPath;
                break;
            case RelationshipTypes.FontTable:
                part = new FontTablePart(this.package as any, normalizedPath);
                partKind = "fontTable";
                this.rolePaths.fontTablePart = normalizedPath;
                break;
            case RelationshipTypes.Numbering:
                part = new NumberingPart(this.package as any, normalizedPath, this.parser);
                partKind = "numbering";
                this.rolePaths.numberingPart = normalizedPath;
                break;
            case RelationshipTypes.Styles:
                this.stylesPart = part = new StylesPart(this.package as any, normalizedPath, this.parser);
                partKind = "styles";
                this.rolePaths.stylesPart = normalizedPath;
                break;
            case RelationshipTypes.Theme:
                part = new ThemePart(this.package as any, normalizedPath);
                partKind = "theme";
                this.rolePaths.themePart = normalizedPath;
                break;
            case RelationshipTypes.Footnotes:
                part = new FootnotesPart(this.package as any, normalizedPath, this.parser);
                partKind = "footnotes";
                this.rolePaths.footnotesPart = normalizedPath;
                break;
            case RelationshipTypes.Endnotes:
                part = new EndnotesPart(this.package as any, normalizedPath, this.parser);
                partKind = "endnotes";
                this.rolePaths.endnotesPart = normalizedPath;
                break;
            case RelationshipTypes.Footer:
                part = new FooterPart(this.package as any, normalizedPath, this.parser);
                partKind = "footer";
                break;
            case RelationshipTypes.Header:
                part = new HeaderPart(this.package as any, normalizedPath, this.parser);
                partKind = "header";
                break;
            case RelationshipTypes.CoreProperties:
                part = new CorePropsPart(this.package as any, normalizedPath);
                partKind = "coreProps";
                this.rolePaths.corePropsPart = normalizedPath;
                break;
            case RelationshipTypes.ExtendedProperties:
                part = new ExtendedPropsPart(this.package as any, normalizedPath);
                partKind = "extendedProps";
                this.rolePaths.extendedPropsPart = normalizedPath;
                break;
            case RelationshipTypes.CustomProperties:
                part = new CustomPropsPart(this.package as any, normalizedPath);
                partKind = "customProps";
                this.rolePaths.customPropsPart = normalizedPath;
                break;
            case RelationshipTypes.Settings:
                part = new SettingsPart(this.package as any, normalizedPath);
                partKind = "settings";
                this.rolePaths.settingsPart = normalizedPath;
                break;
            case RelationshipTypes.Comments:
                part = new CommentsPart(this.package as any, normalizedPath, this.parser);
                partKind = "comments";
                this.rolePaths.commentsPart = normalizedPath;
                break;
            case RelationshipTypes.CommentsExtended:
                part = new CommentsExtendedPart(this.package as any, normalizedPath);
                partKind = "commentsExtended";
                this.rolePaths.commentsExtendedPart = normalizedPath;
                break;
        }

        if (!part)
            return null;

        (part as any).__kind = partKind;
        this.partsMap[normalizedPath] = part;
        this.parts.push(part);

        await part.load();

        if (part.rels?.length > 0) {
            const [folder] = splitPath(part.path);
            await Promise.all(part.rels.map(rel => this.loadRelationshipPart(resolvePath(rel.target, folder), rel.type)));
        }

        return part;
    }

    private collectSnapshotFiles(): SnapshotFile[] {
        const serializedPaths = new Set(this.parts.map(part => normalizePath(part.path)));
        const resourcePaths = new Set<string>();

        for (const part of this.parts) {
            const [folder] = splitPath(part.path);

            for (const rel of part.rels ?? []) {
                if (rel.targetMode === "External")
                    continue;

                const resolvedPath = normalizePath(resolvePath(rel.target, folder));

                if (!this.package.get(resolvedPath) || serializedPaths.has(resolvedPath))
                    continue;

                resourcePaths.add(resolvedPath);
            }
        }

        return Array.from(resourcePaths).map(path => ({
            path,
            buffer: toArrayBuffer(this.package.get(path))
        }));
    }

    private serializePart(part: Part): SerializedPart {
        const base = {
            kind: (part as any).__kind,
            path: part.path,
            rels: part.rels
        };

        switch ((part as any).__kind as PartKind) {
            case "document":
                return {
                    ...base,
                    body: serializeDocumentBody((part as any).body)
                };
            case "fontTable":
                return { ...base, fonts: (part as any).fonts };
            case "numbering":
                return {
                    ...base,
                    numberings: (part as any).numberings,
                    abstractNumberings: (part as any).abstractNumberings,
                    bulletPictures: (part as any).bulletPictures,
                    domNumberings: (part as any).domNumberings
                };
            case "styles":
                return { ...base, styles: (part as any).styles };
            case "theme":
                return { ...base, theme: (part as any).theme };
            case "footnotes":
            case "endnotes":
                return { ...base, notes: (part as any).notes };
            case "header":
            case "footer":
                return { ...base, rootElement: (part as any).rootElement };
            case "coreProps":
            case "extendedProps":
            case "customProps":
                return { ...base, props: (part as any).props };
            case "settings":
                return { ...base, settings: (part as any).settings };
            case "comments":
                return { ...base, comments: (part as any).comments };
            case "commentsExtended":
                return { ...base, comments: (part as any).comments };
            default:
                return null;
        }
    }
}

function normalizeParseOptions(options?: ParseOptions) {
    return {
        breakPages: options?.breakPages ?? defaultOptions.breakPages,
        debug: options?.debug ?? defaultOptions.debug,
        ignoreLastRenderedPageBreak: options?.ignoreLastRenderedPageBreak ?? defaultOptions.ignoreLastRenderedPageBreak,
        inWrapper: options?.inWrapper ?? defaultOptions.inWrapper,
        trimXmlDeclaration: options?.trimXmlDeclaration ?? defaultOptions.trimXmlDeclaration,
    };
}

function normalizeRenderOptions(options?: RenderOptions): Options {
    return {
        ...defaultOptions,
        ...options,
        useWorkerParser: false,
        workerUrl: undefined
    };
}

function validateSnapshot(snapshot: DocxSnapshot) {
    if (!snapshot?.meta)
        throw new Error("DOCX: Invalid snapshot payload");

    if (snapshot.meta.version !== 1)
        throw new Error(`DOCX: Unsupported snapshot version ${snapshot.meta.version}`);
}

function validateRenderOptions(snapshot: DocxSnapshot, options: Options) {
    const parseOptions = snapshot.meta.parseOptions;

    if (options.breakPages !== parseOptions.breakPages) {
        throw new Error("DOCX: renderSnapshot() received breakPages that does not match snapshot parse options");
    }

    if (options.ignoreLastRenderedPageBreak !== parseOptions.ignoreLastRenderedPageBreak) {
        throw new Error("DOCX: renderSnapshot() received ignoreLastRenderedPageBreak that does not match snapshot parse options");
    }
}

function serializeDocumentBody(body: any) {
    if (!body)
        return null;

    return {
        ...body,
        children: []
    };
}

function normalizePath(path: string) {
    return path.startsWith("/") ? path.substring(1) : path;
}

function normalizeFiles(files: Record<string, Uint8Array>) {
    const result: Record<string, Uint8Array> = {};

    for (const [path, file] of Object.entries(files ?? {})) {
        result[normalizePath(path)] = file;
    }

    return result;
}

function toArrayBuffer(data: Uint8Array) {
    return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;
}

async function inputToUint8Array(data: Blob | ArrayBuffer | Uint8Array) {
    if (data instanceof Uint8Array)
        return data.slice();

    if (data instanceof ArrayBuffer)
        return new Uint8Array(data);

    if (ArrayBuffer.isView(data))
        return new Uint8Array(data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength));

    if (typeof data?.arrayBuffer === "function")
        return new Uint8Array(await data.arrayBuffer());

    throw new Error("Unsupported input type for parseToSnapshot");
}

function cloneSerializable<T>(value: T): T {
    if (typeof structuredClone === "function") {
        return structuredClone(value);
    }

    if (Array.isArray(value)) {
        return value.map(item => cloneSerializable(item)) as T;
    }

    if (!value || typeof value !== "object") {
        return value;
    }

    const result: Record<string, any> = {};

    for (const [key, entry] of Object.entries(value)) {
        result[key] = cloneSerializable(entry);
    }

    return result as T;
}

