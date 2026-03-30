import {
	Virtualizer,
	elementScroll,
	observeElementOffset,
	observeElementRect
} from '@tanstack/virtual-core';

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

export class VirtualizedRenderer {
	private virtualizer: Virtualizer<HTMLElement, HTMLElement>;
	private cleanup: () => void = null;
	private elementCache = new Map<number, HTMLElement>();
	private frameId = 0;
	private idleMeasureTimeoutId = 0;
	private idleMeasureFrameId = 0;
	private contentElement: HTMLElement;
	private lastWindowSignature: string = null;
	private isScrolling = false;

	constructor(private options: VirtualizedRendererOptions) {
		this.contentElement = this.createContentElement();
		this.virtualizer = new Virtualizer<HTMLElement, HTMLElement>({
			count: options.items.length,
			getScrollElement: () => options.scrollElement,
			estimateSize: index => options.items[index]?.estimatedSize ?? 0,
			scrollToFn: elementScroll,
			observeElementRect,
			observeElementOffset,
			overscan: options.overscan,
			onChange: (_, sync) => {
				this.isScrolling = sync;
				this.scheduleSync();
				this.scheduleIdleMeasurement(sync ? 120 : 0);
			}
		});
		this.virtualizer.shouldAdjustScrollPositionOnItemSizeChange = () => false;
	}

	mount() {
		if (!this.contentElement.isConnected) {
			this.options.hostElement.replaceChildren(this.contentElement);
		}

		this.cleanup = this.virtualizer._didMount();
		this.virtualizer._willUpdate();
		this.sync();
		this.scheduleIdleMeasurement(0);
	}

	destroy() {
		if (this.frameId) {
			cancelAnimationFrame(this.frameId);
			this.frameId = 0;
		}

		if (this.idleMeasureTimeoutId) {
			clearTimeout(this.idleMeasureTimeoutId);
			this.idleMeasureTimeoutId = 0;
		}

		if (this.idleMeasureFrameId) {
			cancelAnimationFrame(this.idleMeasureFrameId);
			this.idleMeasureFrameId = 0;
		}

		this.cleanup?.();
		this.cleanup = null;
		this.options.hostElement.replaceChildren();
		this.contentElement.replaceChildren();
		this.elementCache.clear();
		this.lastWindowSignature = null;
	}

	get hostElement() {
		return this.options.hostElement;
	}

	private createContentElement() {
		const element = this.options.document.createElement("div");
		element.style.position = "relative";
		element.style.width = "100%";
		element.style.minWidth = "0";
		element.style.boxSizing = "border-box";
		return element;
	}

	private scheduleSync() {
		if (this.frameId) {
			return;
		}

		this.frameId = requestAnimationFrame(() => {
			this.frameId = 0;
			this.sync();
		});
	}

	private scheduleIdleMeasurement(delay: number) {
		if (this.idleMeasureTimeoutId) {
			clearTimeout(this.idleMeasureTimeoutId);
			this.idleMeasureTimeoutId = 0;
		}

		if (this.idleMeasureFrameId) {
			cancelAnimationFrame(this.idleMeasureFrameId);
			this.idleMeasureFrameId = 0;
		}

		this.idleMeasureTimeoutId = window.setTimeout(() => {
			this.idleMeasureTimeoutId = 0;
			this.idleMeasureFrameId = requestAnimationFrame(() => {
				this.idleMeasureFrameId = 0;

				if (this.isScrolling) {
					this.scheduleIdleMeasurement(120);
					return;
				}

				this.measureMountedItems();
			});
		}, delay);
	}

	private sync() {
		const virtualItems = this.virtualizer.getVirtualItems();
		const totalSize = this.virtualizer.getTotalSize();
		const nextIndices = new Set<number>();
		const addedIndices: number[] = [];
		const removedIndices: number[] = [];

		this.contentElement.style.height = `${Math.max(0, totalSize)}px`;

		for (const item of virtualItems) {
			const existing = this.elementCache.get(item.index);
			const element = existing ?? this.options.renderItem(item.index);

			if (!existing) {
				this.prepareItemElement(element);
				this.elementCache.set(item.index, element);
				this.contentElement.appendChild(element);
				addedIndices.push(item.index);
			} else if (!element.isConnected || element.parentElement !== this.contentElement) {
				this.contentElement.appendChild(element);
			}

			this.positionItemElement(element, item.start);
			nextIndices.add(item.index);
		}

		for (const [index, element] of this.elementCache.entries()) {
			if (nextIndices.has(index)) {
				continue;
			}

			element.remove();
			this.elementCache.delete(index);
			removedIndices.push(index);
		}

		this.emitWindowChange(virtualItems, addedIndices, removedIndices);
		this.options.onRendered?.();
	}

	private prepareItemElement(element: HTMLElement) {
		element.style.position = "absolute";
		element.style.left = "0";
		element.style.right = this.options.centerItems ? "0" : "";
		element.style.marginLeft = this.options.centerItems ? "auto" : "";
		element.style.marginRight = this.options.centerItems ? "auto" : "";
		element.style.marginBottom = "0";
		element.style.boxSizing = "border-box";
	}

	private positionItemElement(element: HTMLElement, start: number) {
		element.style.top = `${Math.max(0, Math.round(start))}px`;
	}

	private measureMountedItems() {
		for (const [index, element] of this.elementCache.entries()) {
			if (!element.isConnected) {
				continue;
			}

			const size = Math.ceil(element.getBoundingClientRect().height) + (this.options.itemGap ?? 0);
			this.virtualizer.resizeItem(index, size);
		}
	}

	private emitWindowChange(virtualItems: ReturnType<Virtualizer<HTMLElement, HTMLElement>["getVirtualItems"]>, addedIndices: number[], removedIndices: number[]) {
		const indices = virtualItems.map(item => item.index);
		const signature = indices.join(",");

		if (signature === this.lastWindowSignature) {
			return;
		}

		this.lastWindowSignature = signature;

		if (!indices.length) {
			return;
		}

		this.options.onWindowChange?.({
			startIndex: indices[0],
			endIndex: indices[indices.length - 1],
			indices,
			addedIndices,
			removedIndices,
			items: indices.map(index => ({
				index,
				element: this.elementCache.get(index)
			})),
			isScrolling: this.isScrolling
		});
	}

	getMountedItems() {
		return Array.from(this.elementCache.entries()).map(([index, element]) => ({ index, element }));
	}

	findMountedItem(index: number) {
		return this.elementCache.get(index) ?? null;
	}

	scrollToIndex(index: number, options: {
		block?: ScrollLogicalPosition;
		behavior?: ScrollBehavior;
	} = {}) {
		this.virtualizer.scrollToIndex(index, {
			align: mapBlockToAlign(options.block),
			behavior: options.behavior as any
		});
	}
}

function mapBlockToAlign(block?: ScrollLogicalPosition) {
	switch (block) {
		case "center":
			return "center";
		case "end":
			return "end";
		case "nearest":
			return "auto";
		case "start":
		default:
			return "start";
	}
}
