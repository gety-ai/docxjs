export interface VirtualizedItem {
    key: string | number;
    estimatedSize: number;
}
export interface MountedWindowChange {
    startIndex: number;
    endIndex: number;
    indices: number[];
    addedIndices: number[];
    removedIndices: number[];
    items: Array<{
        index: number;
        element: HTMLElement;
    }>;
    isScrolling: boolean;
}
interface VirtualizedRendererOptions {
    document: Document;
    hostElement: HTMLElement;
    scrollElement: HTMLElement;
    items: VirtualizedItem[];
    overscan: number;
    itemGap?: number;
    centerItems?: boolean;
    renderItem: (index: number) => HTMLElement;
    onRendered?: () => void;
    onWindowChange?: (payload: MountedWindowChange) => void;
}
export declare class VirtualizedRenderer {
    private options;
    private virtualizer;
    private cleanup;
    private elementCache;
    private frameId;
    private idleMeasureTimeoutId;
    private idleMeasureFrameId;
    private contentElement;
    private lastWindowSignature;
    private isScrolling;
    constructor(options: VirtualizedRendererOptions);
    mount(): void;
    destroy(): void;
    get hostElement(): HTMLElement;
    private createContentElement;
    private scheduleSync;
    private scheduleIdleMeasurement;
    private sync;
    private prepareItemElement;
    private positionItemElement;
    private measureMountedItems;
    private emitWindowChange;
    getMountedItems(): {
        index: number;
        element: HTMLElement;
    }[];
    findMountedItem(index: number): HTMLElement;
    scrollToIndex(index: number, options?: {
        block?: ScrollLogicalPosition;
        behavior?: ScrollBehavior;
    }): void;
}
export {};
