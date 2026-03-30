export interface VirtualizedItem {
    key: string | number;
    estimatedSize: number;
}
interface VirtualizedRendererOptions {
    document: Document;
    hostElement: HTMLElement;
    scrollElement: HTMLElement;
    items: VirtualizedItem[];
    overscan: number;
    renderItem: (index: number) => HTMLElement;
    onRendered?: () => void;
}
export declare class VirtualizedRenderer {
    private options;
    private virtualizer;
    private cleanup;
    private elementCache;
    private frameId;
    private topSpacer;
    private bottomSpacer;
    constructor(options: VirtualizedRendererOptions);
    mount(): void;
    destroy(): void;
    get hostElement(): HTMLElement;
    private createSpacer;
    private scheduleSync;
    private sync;
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
