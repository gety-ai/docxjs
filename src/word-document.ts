import { DocumentParser } from './document-parser';
import { Relationship, RelationshipTypes } from './common/relationship';
import { Part } from './common/part';
import { FontTablePart } from './font-table/font-table';
import { OpenXmlPackage, OpenXmlPackageLoadType, OpenXmlPackageSaveType } from './common/open-xml-package';
import { DocumentPart } from './document/document-part';
import { blobToBase64, keyBy, resolvePath, splitPath } from './utils';
import { NumberingPart } from './numbering/numbering-part';
import { StylesPart } from './styles/styles-part';
import { FooterPart, HeaderPart } from "./header-footer/parts";
import { ExtendedPropsPart } from "./document-props/extended-props-part";
import { CorePropsPart } from "./document-props/core-props-part";
import { ThemePart } from "./theme/theme-part";
import { EndnotesPart, FootnotesPart } from "./notes/parts";
import { SettingsPart } from "./settings/settings-part";
import { CustomPropsPart } from "./document-props/custom-props-part";
import { CommentsPart } from "./comments/comments-part";
import { CommentsExtendedPart } from "./comments/comments-extended-part";
import { parseDocumentInWorker, WorkerSessionPackage } from "./worker/client";
import { SerializedWordDocument } from "./worker/worker-types";
import type { DocxSnapshot, SnapshotFile } from "./snapshot";
import type { PaginatedPage } from "./document-pager";

const topLevelRels = [
	{ type: RelationshipTypes.OfficeDocument, target: "word/document.xml" },
	{ type: RelationshipTypes.ExtendedProperties, target: "docProps/app.xml" },
	{ type: RelationshipTypes.CoreProperties, target: "docProps/core.xml" },
	{ type: RelationshipTypes.CustomProperties, target: "docProps/custom.xml" },
];

export class WordDocument {
	private _package: OpenXmlPackage | WorkerSessionPackage;
	private _parser: DocumentParser;
	private _options: any;
	private _objectUrls = new Set<string>();
	private _snapshotPages: PaginatedPage[] = null;

	rels: Relationship[];
	parts: Part[] = [];
	partsMap: Record<string, Part> = {};

	documentPart: DocumentPart;
	fontTablePart: FontTablePart;
	numberingPart: NumberingPart;
	stylesPart: StylesPart;
	footnotesPart: FootnotesPart;
	endnotesPart: EndnotesPart;
	themePart: ThemePart;
	corePropsPart: CorePropsPart;
	extendedPropsPart: ExtendedPropsPart;
	customPropsPart: CustomPropsPart;
	settingsPart: SettingsPart;
	commentsPart: CommentsPart;
	commentsExtendedPart: CommentsExtendedPart;
	pages: PaginatedPage[] = null;

	static async load(blob: Blob | any, parser: DocumentParser, options: any): Promise<WordDocument> {
		var d = new WordDocument();

		d._options = options;
		d._parser = parser;

		if (options.useWorkerParser) {
			try {
				const workerResult = await parseDocumentInWorker(blob, options);

				if (workerResult) {
					d._package = workerResult.package;
					d.applySerializedDocument(workerResult.parsed);
					return d;
				}
			} catch (error) {
				if (options.debug) {
					console.warn("DOCX: Worker parser failed, falling back to main thread", error);
				}
			}
		}

		d._package = await OpenXmlPackage.load(blob, options);
		d.rels = await d._package.loadRelationships();

		await Promise.all(topLevelRels.map(rel => {
			const r = d.rels.find(x => x.type === rel.type) ?? rel; //fallback                    
			return d.loadRelationshipPart(r.target, r.type);
		}));

		return d;
	}

	static fromSnapshot(snapshot: DocxSnapshot, options: any): WordDocument {
		const document = new WordDocument();
		document._options = options;
		document._package = OpenXmlPackage.fromFiles(snapshotFilesToMap(snapshot.files), options);
		document.applySerializedDocument({
			sessionId: "snapshot",
			rels: snapshot.rels,
			parts: materializeSnapshotParts(snapshot.parts),
			rolePaths: snapshot.rolePaths
		} as SerializedWordDocument);
		document.pages = snapshot.pages;
		document._snapshotPages = snapshot.pages;
		return document;
	}

	preparePageForRender(page: PaginatedPage): PaginatedPage {
		return this._snapshotPages ? cloneSerializable(page) : page;
	}

	private applySerializedDocument(data: SerializedWordDocument) {
		this.rels = data.rels;
		this.parts = data.parts as any[];
		this.partsMap = keyBy(this.parts, x => x.path) as any;

		this.documentPart = data.rolePaths.documentPart ? this.partsMap[data.rolePaths.documentPart] as any : null;
		this.fontTablePart = data.rolePaths.fontTablePart ? this.partsMap[data.rolePaths.fontTablePart] as any : null;
		this.numberingPart = data.rolePaths.numberingPart ? this.partsMap[data.rolePaths.numberingPart] as any : null;
		this.stylesPart = data.rolePaths.stylesPart ? this.partsMap[data.rolePaths.stylesPart] as any : null;
		this.footnotesPart = data.rolePaths.footnotesPart ? this.partsMap[data.rolePaths.footnotesPart] as any : null;
		this.endnotesPart = data.rolePaths.endnotesPart ? this.partsMap[data.rolePaths.endnotesPart] as any : null;
		this.themePart = data.rolePaths.themePart ? this.partsMap[data.rolePaths.themePart] as any : null;
		this.corePropsPart = data.rolePaths.corePropsPart ? this.partsMap[data.rolePaths.corePropsPart] as any : null;
		this.extendedPropsPart = data.rolePaths.extendedPropsPart ? this.partsMap[data.rolePaths.extendedPropsPart] as any : null;
		this.customPropsPart = data.rolePaths.customPropsPart ? this.partsMap[data.rolePaths.customPropsPart] as any : null;
		this.settingsPart = data.rolePaths.settingsPart ? this.partsMap[data.rolePaths.settingsPart] as any : null;
		this.commentsPart = data.rolePaths.commentsPart ? this.partsMap[data.rolePaths.commentsPart] as any : null;
		this.commentsExtendedPart = data.rolePaths.commentsExtendedPart ? this.partsMap[data.rolePaths.commentsExtendedPart] as any : null;

		if (this.commentsPart?.comments && !this.commentsPart.commentMap) {
			this.commentsPart.commentMap = keyBy(this.commentsPart.comments, x => x.id) as any;
		}

		if (this.commentsExtendedPart?.comments && !this.commentsExtendedPart.commentMap) {
			this.commentsExtendedPart.commentMap = keyBy(this.commentsExtendedPart.comments, x => x.paraId) as any;
		}
	}

	save(type: OpenXmlPackageSaveType = "blob"): Promise<any> {
		return this._package.save(type);
	}

	async dispose(): Promise<void> {
		for (const url of this._objectUrls) {
			URL.revokeObjectURL(url);
		}

		this._objectUrls.clear();

		if (typeof (this._package as WorkerSessionPackage)?.dispose === "function") {
			await (this._package as WorkerSessionPackage).dispose();
		}
	}

	private async loadRelationshipPart(path: string, type: string): Promise<Part> {
		if (this.partsMap[path])
			return this.partsMap[path];

		if (!this._package.get(path))
			return null;

		let part: Part = null;
		const pkg = this._package as any;

		switch (type) {
			case RelationshipTypes.OfficeDocument:
				this.documentPart = part = new DocumentPart(pkg, path, this._parser);
				break;

			case RelationshipTypes.FontTable:
				this.fontTablePart = part = new FontTablePart(pkg, path);
				break;

			case RelationshipTypes.Numbering:
				this.numberingPart = part = new NumberingPart(pkg, path, this._parser);
				break;

			case RelationshipTypes.Styles:
				this.stylesPart = part = new StylesPart(pkg, path, this._parser);
				break;

			case RelationshipTypes.Theme:
				this.themePart = part = new ThemePart(pkg, path);
				break;

			case RelationshipTypes.Footnotes:
				this.footnotesPart = part = new FootnotesPart(pkg, path, this._parser);
				break;

			case RelationshipTypes.Endnotes:
				this.endnotesPart = part = new EndnotesPart(pkg, path, this._parser);
				break;

			case RelationshipTypes.Footer:
				part = new FooterPart(pkg, path, this._parser);
				break;

			case RelationshipTypes.Header:
				part = new HeaderPart(pkg, path, this._parser);
				break;

			case RelationshipTypes.CoreProperties:
				this.corePropsPart = part = new CorePropsPart(pkg, path);
				break;

			case RelationshipTypes.ExtendedProperties:
				this.extendedPropsPart = part = new ExtendedPropsPart(pkg, path);
				break;

			case RelationshipTypes.CustomProperties:
				part = new CustomPropsPart(pkg, path);
				break;
	
			case RelationshipTypes.Settings:
				this.settingsPart = part = new SettingsPart(pkg, path);
				break;

			case RelationshipTypes.Comments:
				this.commentsPart = part = new CommentsPart(pkg, path, this._parser);
				break;

			case RelationshipTypes.CommentsExtended:
				this.commentsExtendedPart = part = new CommentsExtendedPart(pkg, path);
				break;
		}

		if (part == null)
			return Promise.resolve(null);

		this.partsMap[path] = part;
		this.parts.push(part);

		await part.load();

		if (part.rels?.length > 0) {
			const [folder] = splitPath(part.path);
			await Promise.all(part.rels.map(rel => this.loadRelationshipPart(resolvePath(rel.target, folder), rel.type)));
		}

		return part;
	}

	async loadDocumentImage(id: string, part?: Part): Promise<string> {
		const x = await this.loadResource(part ?? this.documentPart, id, "blob");
		return this.blobToURL(x);
	}

	async loadNumberingImage(id: string): Promise<string> {
		const x = await this.loadResource(this.numberingPart, id, "blob");
		return this.blobToURL(x);
	}

	async loadFont(id: string, key: string): Promise<string> {
		const x = await this.loadResource(this.fontTablePart, id, "uint8array");
		return x ? this.blobToURL(new Blob([deobfuscate(x, key)])) : x;
	}

	async loadAltChunk(id: string, part?: Part): Promise<string> {
		return await this.loadResource(part ?? this.documentPart, id, "string");
	}

	private blobToURL(blob: Blob): string | Promise<string> {
		if (!blob)
			return null;

		if (this._options.useBase64URL) {
			return blobToBase64(blob);
		}

		const url = URL.createObjectURL(blob);
		this._objectUrls.add(url);
		return url;
	}

	findPartByRelId(id: string, basePart: Part = null) {
		var rel = (basePart.rels ?? this.rels).find(r => r.id == id);
		const folder = basePart ? splitPath(basePart.path)[0] : '';
		return rel ? this.partsMap[resolvePath(rel.target, folder)] : null;
	}

	getPathById(part: Part, id: string): string {
		const rel = part.rels.find(x => x.id == id);
		const [folder] = splitPath(part.path);
		return rel ? resolvePath(rel.target, folder) : null;
	}

	private loadResource(part: Part, id: string, outputType: OpenXmlPackageLoadType) {
		const path = this.getPathById(part, id);
		return path ? this._package.load(path, outputType) : Promise.resolve(null);
	}
}

export function deobfuscate(data: Uint8Array, guidKey: string) {
	const len = 16;
	const trimmed = guidKey.replace(/{|}|-/g, "");
	const numbers = new Array(len);

	for (let i = 0; i < len; i++)
		numbers[len - i - 1] = parseInt(trimmed.substring(i * 2, i * 2 + 2), 16);

	for (let i = 0; i < 32; i++)
		data[i] = data[i] ^ numbers[i % len]

	// FIXME: return type
	return data as any;
}

function snapshotFilesToMap(files: SnapshotFile[]) {
	const result: Record<string, Uint8Array> = {};

	for (const file of files ?? []) {
		result[file.path] = new Uint8Array(file.buffer);
	}

	return result;
}

function materializeSnapshotParts(parts: SerializedWordDocument["parts"]) {
	return (parts ?? []).map(part => {
		const materialized = { ...part };

		switch (part.kind) {
			case "styles":
				materialized.styles = cloneSerializable(part.styles);
				break;

			case "header":
			case "footer":
				materialized.rootElement = cloneSerializable(part.rootElement);
				break;

			case "footnotes":
			case "endnotes":
				materialized.notes = cloneSerializable(part.notes);
				break;
		}

		return materialized;
	});
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
