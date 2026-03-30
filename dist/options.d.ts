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
    workerUrl?: string | URL;
    onMountedPageWindowChange?: (payload: {
        startPageIndex: number;
        endPageIndex: number;
        pageIndices: number[];
        addedPageIndices: number[];
        removedPageIndices: number[];
        pages: Array<{
            pageIndex: number;
            element: HTMLElement;
        }>;
        isScrolling: boolean;
    }) => void;
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
