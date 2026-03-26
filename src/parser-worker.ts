import { unzipSync } from "fflate";

import { Relationship, RelationshipTypes, parseRelationships } from "./common/relationship";
import { Part } from "./common/part";
import { OpenXmlPackageOptions } from "./common/open-xml-package";
import { CommentsExtendedPart } from "./comments/comments-extended-part";
import { CommentsPart } from "./comments/comments-part";
import { CorePropsPart } from "./document-props/core-props-part";
import { CustomPropsPart } from "./document-props/custom-props-part";
import { ExtendedPropsPart } from "./document-props/extended-props-part";
import { DocumentPart } from "./document/document-part";
import { DocumentParser } from "./document-parser";
import { FontTablePart } from "./font-table/font-table";
import { FooterPart, HeaderPart } from "./header-footer/parts";
import { EndnotesPart, FootnotesPart } from "./notes/parts";
import { NumberingPart } from "./numbering/numbering-part";
import { XmlParser } from "./parser/xml-parser";
import { parseXmlStringWithTxml } from "./parser/txml-parser";
import { SettingsPart } from "./settings/settings-part";
import { StylesPart } from "./styles/styles-part";
import { ThemePart } from "./theme/theme-part";
import { resolvePath, splitPath } from "./utils";
import { SerializedFileEntry, SerializedPart, SerializedWordDocument } from "./worker/worker-types";

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

class WorkerOpenXmlPackage {
    xmlParser = new XmlParser();
    decoder = new TextDecoder();

    constructor(public files: Record<string, Uint8Array>, public options: OpenXmlPackageOptions) {
    }

    get(path: string) {
        const normalized = normalizePath(path);
        return this.files[normalized] ?? this.files[normalized.replace(/\//g, "\\")] ?? null;
    }

    async load(path: string, type: string = "string"): Promise<any> {
        const file = this.get(path);

        if (!file)
            return null;

        switch (type) {
            case "uint8array":
                return file.slice();
            case "arraybuffer":
                return toArrayBuffer(file);
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

class WorkerDocumentLoader {
    private parser: DocumentParser;
    private package: WorkerOpenXmlPackage;
    private rels: Relationship[] = [];
    private parts: Part[] = [];
    private partsMap: Record<string, Part> = {};
    private rolePaths: Record<string, string> = {};

    constructor(private options: any) {
        this.parser = new DocumentParser(options);
    }

    async load(buffer: ArrayBuffer): Promise<SerializedWordDocument> {
        const files = unzipSync(new Uint8Array(buffer));
        this.package = new WorkerOpenXmlPackage(normalizeFiles(files), this.options);
        this.rels = await this.package.loadRelationships();

        await Promise.all(topLevelRels.map(rel => {
            const relationship = this.rels.find(x => x.type === rel.type) ?? rel;
            return this.loadRelationshipPart(relationship.target, relationship.type);
        }));

        return {
            rels: this.rels,
            parts: this.parts.map(part => this.serializePart(part)).filter(Boolean),
            rolePaths: this.rolePaths,
            files: Object.entries(this.package.files).map(([path, file]) => ({
                path,
                buffer: toArrayBuffer(file)
            }))
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
                part = new DocumentPart(this.package as any, normalizedPath, this.parser);
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
                part = new StylesPart(this.package as any, normalizedPath, this.parser);
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

    private serializePart(part: Part): SerializedPart {
        const base = {
            kind: (part as any).__kind,
            path: part.path,
            rels: part.rels
        };

        switch ((part as any).__kind as PartKind) {
            case "document":
                return { ...base, body: (part as any).body };
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
                return {
                    ...base,
                    comments: (part as any).comments,
                    commentMap: (part as any).commentMap
                };
            case "commentsExtended":
                return {
                    ...base,
                    comments: (part as any).comments,
                    commentMap: (part as any).commentMap
                };
            default:
                return null;
        }
    }
}

self.onmessage = async (event: MessageEvent) => {
    if (event.data?.type !== "parse")
        return;

    try {
        const loader = new WorkerDocumentLoader(event.data.options);
        const payload = await loader.load(event.data.buffer);
        const transfer = payload.files.map((file: SerializedFileEntry) => file.buffer);
        (self as any).postMessage({ type: "parsed", payload }, transfer);
    } catch (error) {
        (self as any).postMessage({
            type: "error",
            error: error?.stack || error?.message || String(error)
        });
    }
};

function normalizeFiles(files: Record<string, Uint8Array>) {
    const result: Record<string, Uint8Array> = {};

    for (const [path, file] of Object.entries(files)) {
        result[normalizePath(path)] = file;
    }

    return result;
}

function normalizePath(path: string) {
    return path.startsWith("/") ? path.substring(1) : path;
}

function toArrayBuffer(file: Uint8Array) {
    return file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength);
}
