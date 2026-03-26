export interface Options {
    inWrapper: boolean;
    hideWrapperOnPrint: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
    trimXmlDeclaration: boolean;
    renderHeaders: boolean;
    renderFooters: boolean;
    renderFootnotes: boolean;
    renderEndnotes: boolean;
    ignoreLastRenderedPageBreak: boolean;
    useBase64URL: boolean;
    renderChanges: boolean;
    renderComments: boolean;
    renderAltChunks: boolean;
    virtualizePages: boolean;
    virtualizePagesOverscan: number;
    useWorkerParser: boolean;
    mergeAdjacent: boolean;
    workerUrl?: string;
    onImageRendered?: (payload: {
        docPrId: string;
        imgEl: HTMLImageElement;
        naturalSize: {
            width: number;
            height: number;
        };
    }) => void;
}
export declare const defaultOptions: Options;
export declare function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<any>;
export declare function renderDocument(document: any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any>;
export declare function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any>;
