import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';
import { defaultOptions, Options } from './options';
export { defaultOptions } from './options';
export type { Options } from './options';
export {
	parseToSnapshot,
	collectSnapshotTransferables,
	renderSnapshot,
	type DocxSnapshot,
	type DocxSnapshotMeta,
	type DocxSnapshotPage,
	type DocxSnapshotSection,
	type ParseOptions,
	type RenderOptions,
	type RenderedSnapshot,
	type SnapshotFile
} from './snapshot';

export function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<any>  {
    const ops = { ...defaultOptions, ...userOptions };
    return WordDocument.load(data, new DocumentParser(ops), ops);
}

export async function renderDocument(document: any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any> {
    const ops = { ...defaultOptions, ...userOptions };
    const renderer = new HtmlRenderer(bodyContainer.ownerDocument ?? window.document);
	return await renderer.render(document, bodyContainer, styleContainer, ops);
}

export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any> {
	const doc = await parseAsync(data, userOptions);
	await renderDocument(doc, bodyContainer, styleContainer, userOptions);
    return doc;
}
