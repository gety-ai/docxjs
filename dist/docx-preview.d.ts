import { Options } from './options';
export { defaultOptions } from './options';
export type { Options } from './options';
export { parseToSnapshot, collectSnapshotTransferables, renderSnapshot, type DocxSnapshot, type DocxSnapshotMeta, type DocxSnapshotPage, type DocxSnapshotSection, type ParseOptions, type RenderOptions, type RenderedSnapshot, type SnapshotFile } from './snapshot';
export declare function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<any>;
export declare function renderDocument(document: any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any>;
export declare function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any>;
